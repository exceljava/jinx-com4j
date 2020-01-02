package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Range extends Com4jObject,Iterable<Com4jObject> {
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
   */

  @DISPID(304)
  java.lang.Object activate();


  /**
   * <p>
   * Getter method for the COM property "AddIndent"
   * </p>
   */

  @DISPID(1063)
  @PropGet
  java.lang.Object getAddIndent();


  /**
   * <p>
   * Setter method for the COM property "AddIndent"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1063)
  @PropPut
  void setAddIndent(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowAbsolute is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnAbsolute is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(236)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddress();

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnAbsolute is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(236)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddress(
    @Optional java.lang.Object rowAbsolute);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, columnAbsolute, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(236)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddress(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, columnAbsolute, referenceStyle, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   */

  @DISPID(236)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddress(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, columnAbsolute, referenceStyle, external, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(236)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddress(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional java.lang.Object external);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   * @param relativeTo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(236)
  @PropGet
  java.lang.String getAddress(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional java.lang.Object external,
    @Optional java.lang.Object relativeTo);


  /**
   * <p>
   * Getter method for the COM property "AddressLocal"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowAbsolute is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnAbsolute is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddressLocal(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(437)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddressLocal();

  /**
   * <p>
   * Getter method for the COM property "AddressLocal"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnAbsolute is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddressLocal(rowAbsolute, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(437)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddressLocal(
    @Optional java.lang.Object rowAbsolute);

  /**
   * <p>
   * Getter method for the COM property "AddressLocal"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddressLocal(rowAbsolute, columnAbsolute, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(437)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddressLocal(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute);

  /**
   * <p>
   * Getter method for the COM property "AddressLocal"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddressLocal(rowAbsolute, columnAbsolute, referenceStyle, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   */

  @DISPID(437)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddressLocal(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle);

  /**
   * <p>
   * Getter method for the COM property "AddressLocal"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddressLocal(rowAbsolute, columnAbsolute, referenceStyle, external, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(437)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.String getAddressLocal(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional java.lang.Object external);

  /**
   * <p>
   * Getter method for the COM property "AddressLocal"
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   * @param relativeTo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(437)
  @PropGet
  java.lang.String getAddressLocal(
    @Optional java.lang.Object rowAbsolute,
    @Optional java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional java.lang.Object external,
    @Optional java.lang.Object relativeTo);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter criteriaRange is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copyToRange is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter unique is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * advancedFilter(action, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param action Mandatory com.exceljava.com4j.excel.XlFilterAction parameter.
   */

  @DISPID(876)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object advancedFilter(
    com.exceljava.com4j.excel.XlFilterAction action);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copyToRange is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter unique is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * advancedFilter(action, criteriaRange, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param action Mandatory com.exceljava.com4j.excel.XlFilterAction parameter.
   * @param criteriaRange Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(876)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object advancedFilter(
    com.exceljava.com4j.excel.XlFilterAction action,
    @Optional java.lang.Object criteriaRange);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter unique is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * advancedFilter(action, criteriaRange, copyToRange, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param action Mandatory com.exceljava.com4j.excel.XlFilterAction parameter.
   * @param criteriaRange Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copyToRange Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(876)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object advancedFilter(
    com.exceljava.com4j.excel.XlFilterAction action,
    @Optional java.lang.Object criteriaRange,
    @Optional java.lang.Object copyToRange);

  /**
   * @param action Mandatory com.exceljava.com4j.excel.XlFilterAction parameter.
   * @param criteriaRange Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copyToRange Optional parameter. Default value is com4j.Variant.getMissing()
   * @param unique Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(876)
  java.lang.Object advancedFilter(
    com.exceljava.com4j.excel.XlFilterAction action,
    @Optional java.lang.Object criteriaRange,
    @Optional java.lang.Object copyToRange,
    @Optional java.lang.Object unique);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter names is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreRelativeAbsolute is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter useRowColumnNames is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter omitColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter omitRow is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlApplyNamesOrder parameter order is set to 1</li><li>java.lang.Object parameter appendLast is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyNames(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(441)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyNames();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreRelativeAbsolute is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter useRowColumnNames is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter omitColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter omitRow is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlApplyNamesOrder parameter order is set to 1</li><li>java.lang.Object parameter appendLast is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyNames(names, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(441)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyNames(
    @Optional java.lang.Object names);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter useRowColumnNames is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter omitColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter omitRow is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlApplyNamesOrder parameter order is set to 1</li><li>java.lang.Object parameter appendLast is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyNames(names, ignoreRelativeAbsolute, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreRelativeAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(441)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyNames(
    @Optional java.lang.Object names,
    @Optional java.lang.Object ignoreRelativeAbsolute);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter omitColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter omitRow is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlApplyNamesOrder parameter order is set to 1</li><li>java.lang.Object parameter appendLast is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyNames(names, ignoreRelativeAbsolute, useRowColumnNames, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreRelativeAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useRowColumnNames Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(441)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyNames(
    @Optional java.lang.Object names,
    @Optional java.lang.Object ignoreRelativeAbsolute,
    @Optional java.lang.Object useRowColumnNames);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter omitRow is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlApplyNamesOrder parameter order is set to 1</li><li>java.lang.Object parameter appendLast is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyNames(names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreRelativeAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useRowColumnNames Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitColumn Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(441)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyNames(
    @Optional java.lang.Object names,
    @Optional java.lang.Object ignoreRelativeAbsolute,
    @Optional java.lang.Object useRowColumnNames,
    @Optional java.lang.Object omitColumn);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlApplyNamesOrder parameter order is set to 1</li><li>java.lang.Object parameter appendLast is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyNames(names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreRelativeAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useRowColumnNames Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitColumn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitRow Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(441)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyNames(
    @Optional java.lang.Object names,
    @Optional java.lang.Object ignoreRelativeAbsolute,
    @Optional java.lang.Object useRowColumnNames,
    @Optional java.lang.Object omitColumn,
    @Optional java.lang.Object omitRow);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter appendLast is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyNames(names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreRelativeAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useRowColumnNames Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitColumn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is 1
   */

  @DISPID(441)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyNames(
    @Optional java.lang.Object names,
    @Optional java.lang.Object ignoreRelativeAbsolute,
    @Optional java.lang.Object useRowColumnNames,
    @Optional java.lang.Object omitColumn,
    @Optional java.lang.Object omitRow,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlApplyNamesOrder order);

  /**
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreRelativeAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useRowColumnNames Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitColumn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is 1
   * @param appendLast Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(441)
  java.lang.Object applyNames(
    @Optional java.lang.Object names,
    @Optional java.lang.Object ignoreRelativeAbsolute,
    @Optional java.lang.Object useRowColumnNames,
    @Optional java.lang.Object omitColumn,
    @Optional java.lang.Object omitRow,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlApplyNamesOrder order,
    @Optional java.lang.Object appendLast);


  /**
   */

  @DISPID(448)
  java.lang.Object applyOutlineStyles();


  /**
   * <p>
   * Getter method for the COM property "Areas"
   * </p>
   */

  @DISPID(568)
  @PropGet
  com.exceljava.com4j.excel.Areas getAreas();


  /**
   * @param string Mandatory java.lang.String parameter.
   */

  @DISPID(1185)
  java.lang.String autoComplete(
    java.lang.String string);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlAutoFillType parameter type is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFill(destination, 0);
   * </code>
   * </p>
   * @param destination Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(449)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlAutoFillType.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(index=-1)
  java.lang.Object autoFill(
    com.exceljava.com4j.excel.Range destination);

  /**
   * @param destination Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param type Optional parameter. Default value is 0
   */

  @DISPID(449)
  java.lang.Object autoFill(
    com.exceljava.com4j.excel.Range destination,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlAutoFillType type);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter field is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter criteria1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlAutoFilterOperator parameter operator is set to 1</li><li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _AutoFilter(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(793)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _AutoFilter();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter criteria1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlAutoFilterOperator parameter operator is set to 1</li><li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _AutoFilter(field, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(793)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _AutoFilter(
    @Optional java.lang.Object field);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlAutoFilterOperator parameter operator is set to 1</li><li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _AutoFilter(field, criteria1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(793)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _AutoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _AutoFilter(field, criteria1, operator, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   */

  @DISPID(793)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _AutoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _AutoFilter(field, criteria1, operator, criteria2, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   * @param criteria2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(793)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _AutoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional java.lang.Object criteria2);

  /**
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   * @param criteria2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visibleDropDown Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(793)
  java.lang.Object _AutoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional java.lang.Object criteria2,
    @Optional java.lang.Object visibleDropDown);


  /**
   */

  @DISPID(237)
  java.lang.Object autoFit();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlRangeAutoFormat parameter format is set to 1</li><li>java.lang.Object parameter number is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter font is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alignment is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter border is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pattern is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(114)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {com.exceljava.com4j.excel.XlRangeAutoFormat.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFormat();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter number is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter font is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alignment is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter border is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pattern is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(format, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param format Optional parameter. Default value is 1
   */

  @DISPID(114)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter font is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alignment is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter border is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pattern is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(format, number, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param format Optional parameter. Default value is 1
   * @param number Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(114)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional java.lang.Object number);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter alignment is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter border is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pattern is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(format, number, font, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param format Optional parameter. Default value is 1
   * @param number Optional parameter. Default value is com4j.Variant.getMissing()
   * @param font Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(114)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional java.lang.Object number,
    @Optional java.lang.Object font);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter border is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pattern is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(format, number, font, alignment, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param format Optional parameter. Default value is 1
   * @param number Optional parameter. Default value is com4j.Variant.getMissing()
   * @param font Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alignment Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(114)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional java.lang.Object number,
    @Optional java.lang.Object font,
    @Optional java.lang.Object alignment);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pattern is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(format, number, font, alignment, border, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param format Optional parameter. Default value is 1
   * @param number Optional parameter. Default value is com4j.Variant.getMissing()
   * @param font Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alignment Optional parameter. Default value is com4j.Variant.getMissing()
   * @param border Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(114)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional java.lang.Object number,
    @Optional java.lang.Object font,
    @Optional java.lang.Object alignment,
    @Optional java.lang.Object border);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(format, number, font, alignment, border, pattern, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param format Optional parameter. Default value is 1
   * @param number Optional parameter. Default value is com4j.Variant.getMissing()
   * @param font Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alignment Optional parameter. Default value is com4j.Variant.getMissing()
   * @param border Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pattern Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(114)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional java.lang.Object number,
    @Optional java.lang.Object font,
    @Optional java.lang.Object alignment,
    @Optional java.lang.Object border,
    @Optional java.lang.Object pattern);

  /**
   * @param format Optional parameter. Default value is 1
   * @param number Optional parameter. Default value is com4j.Variant.getMissing()
   * @param font Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alignment Optional parameter. Default value is com4j.Variant.getMissing()
   * @param border Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pattern Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(114)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional java.lang.Object number,
    @Optional java.lang.Object font,
    @Optional java.lang.Object alignment,
    @Optional java.lang.Object border,
    @Optional java.lang.Object pattern,
    @Optional java.lang.Object width);


  /**
   */

  @DISPID(1036)
  java.lang.Object autoOutline();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lineStyle is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlBorderWeight parameter weight is set to 2</li><li>com.exceljava.com4j.excel.XlColorIndex parameter colorIndex is set to -4105</li><li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _BorderAround(com4j.Variant.getMissing(), 2, -4105, com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1067)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "2", "-4105", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _BorderAround();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlBorderWeight parameter weight is set to 2</li><li>com.exceljava.com4j.excel.XlColorIndex parameter colorIndex is set to -4105</li><li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _BorderAround(lineStyle, 2, -4105, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1067)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"2", "-4105", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _BorderAround(
    @Optional java.lang.Object lineStyle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlColorIndex parameter colorIndex is set to -4105</li><li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _BorderAround(lineStyle, weight, -4105, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   */

  @DISPID(1067)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"-4105", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _BorderAround(
    @Optional java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _BorderAround(lineStyle, weight, colorIndex, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   * @param colorIndex Optional parameter. Default value is -4105
   */

  @DISPID(1067)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _BorderAround(
    @Optional java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex);

  /**
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   * @param colorIndex Optional parameter. Default value is -4105
   * @param color Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1067)
  java.lang.Object _BorderAround(
    @Optional java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex,
    @Optional java.lang.Object color);


  /**
   * <p>
   * Getter method for the COM property "Borders"
   * </p>
   */

  @DISPID(435)
  @PropGet
  com.exceljava.com4j.excel.Borders getBorders();


  /**
   */

  @DISPID(279)
  java.lang.Object calculate();


  /**
   * <p>
   * Getter method for the COM property "Cells"
   * </p>
   */

  @DISPID(238)
  @PropGet
  com.exceljava.com4j.excel.Range getCells();


  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(603)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Characters getCharacters();

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(start, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(603)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional java.lang.Object start);

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(603)
  @PropGet
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional java.lang.Object start,
    @Optional java.lang.Object length);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter customDictionary is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(505)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object checkSpelling();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object checkSpelling(
    @Optional java.lang.Object customDictionary);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, ignoreUppercase, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object checkSpelling(
    @Optional java.lang.Object customDictionary,
    @Optional java.lang.Object ignoreUppercase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, ignoreUppercase, alwaysSuggest, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object checkSpelling(
    @Optional java.lang.Object customDictionary,
    @Optional java.lang.Object ignoreUppercase,
    @Optional java.lang.Object alwaysSuggest);

  /**
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505)
  java.lang.Object checkSpelling(
    @Optional java.lang.Object customDictionary,
    @Optional java.lang.Object ignoreUppercase,
    @Optional java.lang.Object alwaysSuggest,
    @Optional java.lang.Object spellLang);


  /**
   */

  @DISPID(111)
  java.lang.Object clear();


  /**
   */

  @DISPID(113)
  java.lang.Object clearContents();


  /**
   */

  @DISPID(112)
  java.lang.Object clearFormats();


  /**
   */

  @DISPID(239)
  java.lang.Object clearNotes();


  /**
   */

  @DISPID(1037)
  java.lang.Object clearOutline();


  /**
   * <p>
   * Getter method for the COM property "Column"
   * </p>
   */

  @DISPID(240)
  @PropGet
  int getColumn();


  /**
   * @param comparison Mandatory java.lang.Object parameter.
   */

  @DISPID(510)
  com.exceljava.com4j.excel.Range columnDifferences(
    java.lang.Object comparison);


  /**
   * <p>
   * Getter method for the COM property "Columns"
   * </p>
   */

  @DISPID(241)
  @PropGet
  com.exceljava.com4j.excel.Range getColumns();


  /**
   * <p>
   * Getter method for the COM property "ColumnWidth"
   * </p>
   */

  @DISPID(242)
  @PropGet
  java.lang.Object getColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "ColumnWidth"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(242)
  @PropPut
  void setColumnWidth(
    java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sources is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter function is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter topRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter leftColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createLinks is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * consolidate(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(482)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object consolidate();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter function is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter topRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter leftColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createLinks is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * consolidate(sources, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sources Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(482)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object consolidate(
    @Optional java.lang.Object sources);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter topRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter leftColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createLinks is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * consolidate(sources, function, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sources Optional parameter. Default value is com4j.Variant.getMissing()
   * @param function Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(482)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object consolidate(
    @Optional java.lang.Object sources,
    @Optional java.lang.Object function);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter leftColumn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createLinks is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * consolidate(sources, function, topRow, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sources Optional parameter. Default value is com4j.Variant.getMissing()
   * @param function Optional parameter. Default value is com4j.Variant.getMissing()
   * @param topRow Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(482)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object consolidate(
    @Optional java.lang.Object sources,
    @Optional java.lang.Object function,
    @Optional java.lang.Object topRow);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createLinks is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * consolidate(sources, function, topRow, leftColumn, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sources Optional parameter. Default value is com4j.Variant.getMissing()
   * @param function Optional parameter. Default value is com4j.Variant.getMissing()
   * @param topRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param leftColumn Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(482)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object consolidate(
    @Optional java.lang.Object sources,
    @Optional java.lang.Object function,
    @Optional java.lang.Object topRow,
    @Optional java.lang.Object leftColumn);

  /**
   * @param sources Optional parameter. Default value is com4j.Variant.getMissing()
   * @param function Optional parameter. Default value is com4j.Variant.getMissing()
   * @param topRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param leftColumn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createLinks Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(482)
  java.lang.Object consolidate(
    @Optional java.lang.Object sources,
    @Optional java.lang.Object function,
    @Optional java.lang.Object topRow,
    @Optional java.lang.Object leftColumn,
    @Optional java.lang.Object createLinks);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copy(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(551)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object copy();

  /**
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(551)
  java.lang.Object copy(
    @Optional java.lang.Object destination);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter maxRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter maxColumns is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyFromRecordset(data, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param data Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1152)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  int copyFromRecordset(
    com4j.Com4jObject data);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter maxColumns is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyFromRecordset(data, maxRows, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param data Mandatory com4j.Com4jObject parameter.
   * @param maxRows Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1152)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  int copyFromRecordset(
    com4j.Com4jObject data,
    @Optional java.lang.Object maxRows);

  /**
   * @param data Mandatory com4j.Com4jObject parameter.
   * @param maxRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param maxColumns Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1152)
  int copyFromRecordset(
    com4j.Com4jObject data,
    @Optional java.lang.Object maxRows,
    @Optional java.lang.Object maxColumns);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(1, -4147);
   * </code>
   * </p>
   */

  @DISPID(213)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "-4147"})
  @ReturnValue(index=-1)
  java.lang.Object copyPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, -4147);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 1
   */

  @DISPID(213)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-4147"})
  @ReturnValue(index=-1)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * @param appearance Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is -4147
   */

  @DISPID(213)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter bottom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter right is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createNames(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(457)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createNames();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter bottom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter right is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createNames(top, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(457)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createNames(
    @Optional java.lang.Object top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter bottom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter right is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createNames(top, left, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(457)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createNames(
    @Optional java.lang.Object top,
    @Optional java.lang.Object left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter right is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createNames(top, left, bottom, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param bottom Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(457)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createNames(
    @Optional java.lang.Object top,
    @Optional java.lang.Object left,
    @Optional java.lang.Object bottom);

  /**
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param bottom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param right Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(457)
  java.lang.Object createNames(
    @Optional java.lang.Object top,
    @Optional java.lang.Object left,
    @Optional java.lang.Object bottom,
    @Optional java.lang.Object right);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter edition is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>java.lang.Object parameter containsPICT is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(458)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createPublisher();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>java.lang.Object parameter containsPICT is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createPublisher(
    @Optional java.lang.Object edition);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsPICT is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   */

  @DISPID(458)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createPublisher(
    @Optional java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, containsPICT, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createPublisher(
    @Optional java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional java.lang.Object containsPICT);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, containsPICT, containsBIFF, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createPublisher(
    @Optional java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional java.lang.Object containsPICT,
    @Optional java.lang.Object containsBIFF);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, containsPICT, containsBIFF, containsRTF, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsRTF Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createPublisher(
    @Optional java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional java.lang.Object containsPICT,
    @Optional java.lang.Object containsBIFF,
    @Optional java.lang.Object containsRTF);

  /**
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsRTF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsVALU Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458)
  java.lang.Object createPublisher(
    @Optional java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional java.lang.Object containsPICT,
    @Optional java.lang.Object containsBIFF,
    @Optional java.lang.Object containsRTF,
    @Optional java.lang.Object containsVALU);


  /**
   * <p>
   * Getter method for the COM property "CurrentArray"
   * </p>
   */

  @DISPID(501)
  @PropGet
  com.exceljava.com4j.excel.Range getCurrentArray();


  /**
   * <p>
   * Getter method for the COM property "CurrentRegion"
   * </p>
   */

  @DISPID(243)
  @PropGet
  com.exceljava.com4j.excel.Range getCurrentRegion();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * cut(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(565)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object cut();

  /**
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(565)
  java.lang.Object cut(
    @Optional java.lang.Object destination);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowcol is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlDataSeriesType parameter type is set to -4132</li><li>com.exceljava.com4j.excel.XlDataSeriesDate parameter date is set to 1</li><li>java.lang.Object parameter step is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter stop is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trend is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dataSeries(com4j.Variant.getMissing(), -4132, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(464)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlDataSeriesType.class, com.exceljava.com4j.excel.XlDataSeriesDate.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "-4132", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object dataSeries();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlDataSeriesType parameter type is set to -4132</li><li>com.exceljava.com4j.excel.XlDataSeriesDate parameter date is set to 1</li><li>java.lang.Object parameter step is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter stop is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trend is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dataSeries(rowcol, -4132, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(464)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlDataSeriesType.class, com.exceljava.com4j.excel.XlDataSeriesDate.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4132", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object dataSeries(
    @Optional java.lang.Object rowcol);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlDataSeriesDate parameter date is set to 1</li><li>java.lang.Object parameter step is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter stop is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trend is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dataSeries(rowcol, type, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is -4132
   */

  @DISPID(464)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlDataSeriesDate.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object dataSeries(
    @Optional java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter step is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter stop is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trend is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dataSeries(rowcol, type, date, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is -4132
   * @param date Optional parameter. Default value is 1
   */

  @DISPID(464)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object dataSeries(
    @Optional java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlDataSeriesDate date);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter stop is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trend is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dataSeries(rowcol, type, date, step, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is -4132
   * @param date Optional parameter. Default value is 1
   * @param step Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(464)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object dataSeries(
    @Optional java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlDataSeriesDate date,
    @Optional java.lang.Object step);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter trend is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dataSeries(rowcol, type, date, step, stop, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is -4132
   * @param date Optional parameter. Default value is 1
   * @param step Optional parameter. Default value is com4j.Variant.getMissing()
   * @param stop Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(464)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object dataSeries(
    @Optional java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlDataSeriesDate date,
    @Optional java.lang.Object step,
    @Optional java.lang.Object stop);

  /**
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is -4132
   * @param date Optional parameter. Default value is 1
   * @param step Optional parameter. Default value is com4j.Variant.getMissing()
   * @param stop Optional parameter. Default value is com4j.Variant.getMissing()
   * @param trend Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(464)
  java.lang.Object dataSeries(
    @Optional java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlDataSeriesDate date,
    @Optional java.lang.Object step,
    @Optional java.lang.Object stop,
    @Optional java.lang.Object trend);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object get_Default();

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(rowIndex, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object get_Default(
    @Optional java.lang.Object rowIndex);

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  java.lang.Object get_Default(
    @Optional java.lang.Object rowIndex,
    @Optional java.lang.Object columnIndex);


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * set_Default(com4j.Variant.getMissing(), com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropPut
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void set_Default(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * set_Default(rowIndex, com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropPut
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void set_Default(
    @Optional java.lang.Object rowIndex,
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropPut
  @DefaultMethod
  void set_Default(
    @Optional java.lang.Object rowIndex,
    @Optional java.lang.Object columnIndex,
    java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter shift is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * delete(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(117)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object delete();

  /**
   * @param shift Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(117)
  java.lang.Object delete(
    @Optional java.lang.Object shift);


  /**
   * <p>
   * Getter method for the COM property "Dependents"
   * </p>
   */

  @DISPID(543)
  @PropGet
  com.exceljava.com4j.excel.Range getDependents();


  /**
   */

  @DISPID(245)
  java.lang.Object dialogBox();


  /**
   * <p>
   * Getter method for the COM property "DirectDependents"
   * </p>
   */

  @DISPID(545)
  @PropGet
  com.exceljava.com4j.excel.Range getDirectDependents();


  /**
   * <p>
   * Getter method for the COM property "DirectPrecedents"
   * </p>
   */

  @DISPID(546)
  @PropGet
  com.exceljava.com4j.excel.Range getDirectPrecedents();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reference is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter chartSize is set to 1</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * editionOptions(type, option, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlEditionType parameter.
   * @param option Mandatory com.exceljava.com4j.excel.XlEditionOptionsOption parameter.
   */

  @DISPID(1131)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter reference is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter chartSize is set to 1</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * editionOptions(type, option, name, com4j.Variant.getMissing(), 1, 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlEditionType parameter.
   * @param option Mandatory com.exceljava.com4j.excel.XlEditionOptionsOption parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1131)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional java.lang.Object name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter chartSize is set to 1</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * editionOptions(type, option, name, reference, 1, 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlEditionType parameter.
   * @param option Mandatory com.exceljava.com4j.excel.XlEditionOptionsOption parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reference Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1131)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional java.lang.Object name,
    @Optional java.lang.Object reference);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter chartSize is set to 1</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * editionOptions(type, option, name, reference, appearance, 1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlEditionType parameter.
   * @param option Mandatory com.exceljava.com4j.excel.XlEditionOptionsOption parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reference Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   */

  @DISPID(1131)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional java.lang.Object name,
    @Optional java.lang.Object reference,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * editionOptions(type, option, name, reference, appearance, chartSize, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlEditionType parameter.
   * @param option Mandatory com.exceljava.com4j.excel.XlEditionOptionsOption parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reference Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param chartSize Optional parameter. Default value is 1
   */

  @DISPID(1131)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional java.lang.Object name,
    @Optional java.lang.Object reference,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance chartSize);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlEditionType parameter.
   * @param option Mandatory com.exceljava.com4j.excel.XlEditionOptionsOption parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reference Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param chartSize Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1131)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional java.lang.Object name,
    @Optional java.lang.Object reference,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance chartSize,
    @Optional java.lang.Object format);


  /**
   * <p>
   * Getter method for the COM property "End"
   * </p>
   * @param direction Mandatory com.exceljava.com4j.excel.XlDirection parameter.
   */

  @DISPID(500)
  @PropGet
  com.exceljava.com4j.excel.Range getEnd(
    com.exceljava.com4j.excel.XlDirection direction);


  /**
   * <p>
   * Getter method for the COM property "EntireColumn"
   * </p>
   */

  @DISPID(246)
  @PropGet
  com.exceljava.com4j.excel.Range getEntireColumn();


  /**
   * <p>
   * Getter method for the COM property "EntireRow"
   * </p>
   */

  @DISPID(247)
  @PropGet
  com.exceljava.com4j.excel.Range getEntireRow();


  /**
   */

  @DISPID(248)
  java.lang.Object fillDown();


  /**
   */

  @DISPID(249)
  java.lang.Object fillLeft();


  /**
   */

  @DISPID(250)
  java.lang.Object fillRight();


  /**
   */

  @DISPID(251)
  java.lang.Object fillUp();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter after is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter lookIn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter lookAt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSearchDirection parameter searchDirection is set to 1</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lookIn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter lookAt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSearchDirection parameter searchDirection is set to 1</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, after, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lookAt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSearchDirection parameter searchDirection is set to 1</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, after, lookIn, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookIn Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after,
    @Optional java.lang.Object lookIn);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSearchDirection parameter searchDirection is set to 1</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, after, lookIn, lookAt, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookIn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after,
    @Optional java.lang.Object lookIn,
    @Optional java.lang.Object lookAt);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSearchDirection parameter searchDirection is set to 1</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, after, lookIn, lookAt, searchOrder, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookIn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after,
    @Optional java.lang.Object lookIn,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, after, lookIn, lookAt, searchOrder, searchDirection, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookIn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchDirection Optional parameter. Default value is 1
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after,
    @Optional java.lang.Object lookIn,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSearchDirection searchDirection);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookIn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchDirection Optional parameter. Default value is 1
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after,
    @Optional java.lang.Object lookIn,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSearchDirection searchDirection,
    @Optional java.lang.Object matchCase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookIn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchDirection Optional parameter. Default value is 1
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(398)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after,
    @Optional java.lang.Object lookIn,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSearchDirection searchDirection,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte);

  /**
   * @param what Mandatory java.lang.Object parameter.
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookIn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchDirection Optional parameter. Default value is 1
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(398)
  com.exceljava.com4j.excel.Range find(
    java.lang.Object what,
    @Optional java.lang.Object after,
    @Optional java.lang.Object lookIn,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSearchDirection searchDirection,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte,
    @Optional java.lang.Object searchFormat);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter after is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * findNext(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(399)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range findNext();

  /**
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(399)
  com.exceljava.com4j.excel.Range findNext(
    @Optional java.lang.Object after);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter after is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * findPrevious(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(400)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range findPrevious();

  /**
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(400)
  com.exceljava.com4j.excel.Range findPrevious(
    @Optional java.lang.Object after);


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   */

  @DISPID(146)
  @PropGet
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   */

  @DISPID(261)
  @PropGet
  java.lang.Object getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(261)
  @PropPut
  void setFormula(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaArray"
   * </p>
   */

  @DISPID(586)
  @PropGet
  java.lang.Object getFormulaArray();


  /**
   * <p>
   * Setter method for the COM property "FormulaArray"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(586)
  @PropPut
  void setFormulaArray(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaLabel"
   * </p>
   */

  @DISPID(1380)
  @PropGet
  com.exceljava.com4j.excel.XlFormulaLabel getFormulaLabel();


  /**
   * <p>
   * Setter method for the COM property "FormulaLabel"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlFormulaLabel parameter.
   */

  @DISPID(1380)
  @PropPut
  void setFormulaLabel(
    com.exceljava.com4j.excel.XlFormulaLabel rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaHidden"
   * </p>
   */

  @DISPID(262)
  @PropGet
  java.lang.Object getFormulaHidden();


  /**
   * <p>
   * Setter method for the COM property "FormulaHidden"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(262)
  @PropPut
  void setFormulaHidden(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaLocal"
   * </p>
   */

  @DISPID(263)
  @PropGet
  java.lang.Object getFormulaLocal();


  /**
   * <p>
   * Setter method for the COM property "FormulaLocal"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(263)
  @PropPut
  void setFormulaLocal(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaR1C1"
   * </p>
   */

  @DISPID(264)
  @PropGet
  java.lang.Object getFormulaR1C1();


  /**
   * <p>
   * Setter method for the COM property "FormulaR1C1"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(264)
  @PropPut
  void setFormulaR1C1(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaR1C1Local"
   * </p>
   */

  @DISPID(265)
  @PropGet
  java.lang.Object getFormulaR1C1Local();


  /**
   * <p>
   * Setter method for the COM property "FormulaR1C1Local"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(265)
  @PropPut
  void setFormulaR1C1Local(
    java.lang.Object rhs);


  /**
   */

  @DISPID(571)
  java.lang.Object functionWizard();


  /**
   * @param goal Mandatory java.lang.Object parameter.
   * @param changingCell Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(472)
  boolean goalSeek(
    java.lang.Object goal,
    com.exceljava.com4j.excel.Range changingCell);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter end is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter by is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter periods is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * group(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(46)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object group();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter end is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter by is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter periods is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * group(start, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(46)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object group(
    @Optional java.lang.Object start);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter by is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter periods is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * group(start, end, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param end Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(46)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object group(
    @Optional java.lang.Object start,
    @Optional java.lang.Object end);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter periods is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * group(start, end, by, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param end Optional parameter. Default value is com4j.Variant.getMissing()
   * @param by Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(46)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object group(
    @Optional java.lang.Object start,
    @Optional java.lang.Object end,
    @Optional java.lang.Object by);

  /**
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param end Optional parameter. Default value is com4j.Variant.getMissing()
   * @param by Optional parameter. Default value is com4j.Variant.getMissing()
   * @param periods Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(46)
  java.lang.Object group(
    @Optional java.lang.Object start,
    @Optional java.lang.Object end,
    @Optional java.lang.Object by,
    @Optional java.lang.Object periods);


  /**
   * <p>
   * Getter method for the COM property "HasArray"
   * </p>
   */

  @DISPID(266)
  @PropGet
  java.lang.Object getHasArray();


  /**
   * <p>
   * Getter method for the COM property "HasFormula"
   * </p>
   */

  @DISPID(267)
  @PropGet
  java.lang.Object getHasFormula();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   */

  @DISPID(123)
  @PropGet
  java.lang.Object getHeight();


  /**
   * <p>
   * Getter method for the COM property "Hidden"
   * </p>
   */

  @DISPID(268)
  @PropGet
  java.lang.Object getHidden();


  /**
   * <p>
   * Setter method for the COM property "Hidden"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(268)
  @PropPut
  void setHidden(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalAlignment"
   * </p>
   */

  @DISPID(136)
  @PropGet
  java.lang.Object getHorizontalAlignment();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAlignment"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(136)
  @PropPut
  void setHorizontalAlignment(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "IndentLevel"
   * </p>
   */

  @DISPID(201)
  @PropGet
  java.lang.Object getIndentLevel();


  /**
   * <p>
   * Setter method for the COM property "IndentLevel"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(201)
  @PropPut
  void setIndentLevel(
    java.lang.Object rhs);


  /**
   * @param insertAmount Mandatory int parameter.
   */

  @DISPID(1381)
  void insertIndent(
    int insertAmount);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter shift is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copyOrigin is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(252)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object insert();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copyOrigin is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert(shift, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param shift Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(252)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object insert(
    @Optional java.lang.Object shift);

  /**
   * @param shift Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copyOrigin Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(252)
  java.lang.Object insert(
    @Optional java.lang.Object shift,
    @Optional java.lang.Object copyOrigin);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   */

  @DISPID(129)
  @PropGet
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getItem(rowIndex, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getItem(
    java.lang.Object rowIndex);

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(170)
  @PropGet
  java.lang.Object getItem(
    java.lang.Object rowIndex,
    @Optional java.lang.Object columnIndex);


  /**
   * <p>
   * Setter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setItem(rowIndex, com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setItem(
    java.lang.Object rowIndex,
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Item"
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropPut
  void setItem(
    java.lang.Object rowIndex,
    @Optional java.lang.Object columnIndex,
    java.lang.Object rhs);


  /**
   */

  @DISPID(495)
  java.lang.Object justify();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   */

  @DISPID(127)
  @PropGet
  java.lang.Object getLeft();


  /**
   * <p>
   * Getter method for the COM property "ListHeaderRows"
   * </p>
   */

  @DISPID(1187)
  @PropGet
  int getListHeaderRows();


  /**
   */

  @DISPID(253)
  java.lang.Object listNames();


  /**
   * <p>
   * Getter method for the COM property "LocationInTable"
   * </p>
   */

  @DISPID(691)
  @PropGet
  com.exceljava.com4j.excel.XlLocationInTable getLocationInTable();


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   */

  @DISPID(269)
  @PropGet
  java.lang.Object getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(269)
  @PropPut
  void setLocked(
    java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter across is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * merge(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(564)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void merge();

  /**
   * @param across Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(564)
  void merge(
    @Optional java.lang.Object across);


  /**
   */

  @DISPID(1384)
  void unMerge();


  /**
   * <p>
   * Getter method for the COM property "MergeArea"
   * </p>
   */

  @DISPID(1385)
  @PropGet
  com.exceljava.com4j.excel.Range getMergeArea();


  /**
   * <p>
   * Getter method for the COM property "MergeCells"
   * </p>
   */

  @DISPID(208)
  @PropGet
  java.lang.Object getMergeCells();


  /**
   * <p>
   * Setter method for the COM property "MergeCells"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(208)
  @PropPut
  void setMergeCells(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.Object getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(110)
  @PropPut
  void setName(
    java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter towardPrecedent is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arrowNumber is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkNumber is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * navigateArrow(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1032)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object navigateArrow();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arrowNumber is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkNumber is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * navigateArrow(towardPrecedent, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param towardPrecedent Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1032)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object navigateArrow(
    @Optional java.lang.Object towardPrecedent);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter linkNumber is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * navigateArrow(towardPrecedent, arrowNumber, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param towardPrecedent Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arrowNumber Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1032)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object navigateArrow(
    @Optional java.lang.Object towardPrecedent,
    @Optional java.lang.Object arrowNumber);

  /**
   * @param towardPrecedent Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arrowNumber Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkNumber Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1032)
  java.lang.Object navigateArrow(
    @Optional java.lang.Object towardPrecedent,
    @Optional java.lang.Object arrowNumber,
    @Optional java.lang.Object linkNumber);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Next"
   * </p>
   */

  @DISPID(502)
  @PropGet
  com.exceljava.com4j.excel.Range getNext();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter text is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * noteText(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1127)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String noteText();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * noteText(text, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1127)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String noteText(
    @Optional java.lang.Object text);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * noteText(text, start, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1127)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.String noteText(
    @Optional java.lang.Object text,
    @Optional java.lang.Object start);

  /**
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1127)
  java.lang.String noteText(
    @Optional java.lang.Object text,
    @Optional java.lang.Object start,
    @Optional java.lang.Object length);


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   */

  @DISPID(193)
  @PropGet
  java.lang.Object getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(193)
  @PropPut
  void setNumberFormat(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormatLocal"
   * </p>
   */

  @DISPID(1097)
  @PropGet
  java.lang.Object getNumberFormatLocal();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLocal"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1097)
  @PropPut
  void setNumberFormatLocal(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Offset"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowOffset is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnOffset is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOffset(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(254)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range getOffset();

  /**
   * <p>
   * Getter method for the COM property "Offset"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnOffset is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOffset(rowOffset, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowOffset Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(254)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range getOffset(
    @Optional java.lang.Object rowOffset);

  /**
   * <p>
   * Getter method for the COM property "Offset"
   * </p>
   * @param rowOffset Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnOffset Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(254)
  @PropGet
  com.exceljava.com4j.excel.Range getOffset(
    @Optional java.lang.Object rowOffset,
    @Optional java.lang.Object columnOffset);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   */

  @DISPID(134)
  @PropGet
  java.lang.Object getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(134)
  @PropPut
  void setOrientation(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "OutlineLevel"
   * </p>
   */

  @DISPID(271)
  @PropGet
  java.lang.Object getOutlineLevel();


  /**
   * <p>
   * Setter method for the COM property "OutlineLevel"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(271)
  @PropPut
  void setOutlineLevel(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PageBreak"
   * </p>
   */

  @DISPID(255)
  @PropGet
  int getPageBreak();


  /**
   * <p>
   * Setter method for the COM property "PageBreak"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(255)
  @PropPut
  void setPageBreak(
    int rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter parseLine is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * parse(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(477)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object parse();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * parse(parseLine, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param parseLine Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(477)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object parse(
    @Optional java.lang.Object parseLine);

  /**
   * @param parseLine Optional parameter. Default value is com4j.Variant.getMissing()
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(477)
  java.lang.Object parse(
    @Optional java.lang.Object parseLine,
    @Optional java.lang.Object destination);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPasteType parameter paste is set to -4104</li><li>com.exceljava.com4j.excel.XlPasteSpecialOperation parameter operation is set to -4142</li><li>java.lang.Object parameter skipBlanks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(-4104, -4142, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1027)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteType.class, com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4104", "-4142", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PasteSpecial();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPasteSpecialOperation parameter operation is set to -4142</li><li>java.lang.Object parameter skipBlanks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(paste, -4142, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param paste Optional parameter. Default value is -4104
   */

  @DISPID(1027)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4142", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter skipBlanks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(paste, operation, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   */

  @DISPID(1027)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(paste, operation, skipBlanks, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   * @param skipBlanks Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional java.lang.Object skipBlanks);

  /**
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   * @param skipBlanks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param transpose Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027)
  java.lang.Object _PasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional java.lang.Object skipBlanks,
    @Optional java.lang.Object transpose);


  /**
   * <p>
   * Getter method for the COM property "PivotField"
   * </p>
   */

  @DISPID(731)
  @PropGet
  com.exceljava.com4j.excel.PivotField getPivotField();


  /**
   * <p>
   * Getter method for the COM property "PivotItem"
   * </p>
   */

  @DISPID(740)
  @PropGet
  com.exceljava.com4j.excel.PivotItem getPivotItem();


  /**
   * <p>
   * Getter method for the COM property "PivotTable"
   * </p>
   */

  @DISPID(716)
  @PropGet
  com.exceljava.com4j.excel.PivotTable getPivotTable();


  /**
   * <p>
   * Getter method for the COM property "Precedents"
   * </p>
   */

  @DISPID(544)
  @PropGet
  com.exceljava.com4j.excel.Range getPrecedents();


  /**
   * <p>
   * Getter method for the COM property "PrefixCharacter"
   * </p>
   */

  @DISPID(504)
  @PropGet
  java.lang.Object getPrefixCharacter();


  /**
   * <p>
   * Getter method for the COM property "Previous"
   * </p>
   */

  @DISPID(503)
  @PropGet
  com.exceljava.com4j.excel.Range getPrevious();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(905)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object __PrintOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object __PrintOut(
    @Optional java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object __PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object __PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object __PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object __PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object __PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905)
  java.lang.Object __PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile,
    @Optional java.lang.Object collate);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter enableChanges is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printPreview(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(281)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printPreview();

  /**
   * @param enableChanges Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(281)
  java.lang.Object printPreview(
    @Optional java.lang.Object enableChanges);


  /**
   * <p>
   * Getter method for the COM property "QueryTable"
   * </p>
   */

  @DISPID(1386)
  @PropGet
  com.exceljava.com4j.excel._QueryTable getQueryTable();


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
   */

  @DISPID(197)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range getRange(
    java.lang.Object cell1);

  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * @param cell1 Mandatory java.lang.Object parameter.
   * @param cell2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(197)
  @PropGet
  com.exceljava.com4j.excel.Range getRange(
    java.lang.Object cell1,
    @Optional java.lang.Object cell2);


  /**
   */

  @DISPID(883)
  java.lang.Object removeSubtotal();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lookAt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Replace(what, replacement, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   */

  @DISPID(226)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean _Replace(
    java.lang.Object what,
    java.lang.Object replacement);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Replace(what, replacement, lookAt, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(226)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean _Replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Replace(what, replacement, lookAt, searchOrder, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(226)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean _Replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Replace(what, replacement, lookAt, searchOrder, matchCase, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(226)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean _Replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Replace(what, replacement, lookAt, searchOrder, matchCase, matchByte, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(226)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean _Replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Replace(what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(226)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  boolean _Replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte,
    @Optional java.lang.Object searchFormat);

  /**
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replaceFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(226)
  boolean _Replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte,
    @Optional java.lang.Object searchFormat,
    @Optional java.lang.Object replaceFormat);


  /**
   * <p>
   * Getter method for the COM property "Resize"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnSize is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getResize(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(256)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range getResize();

  /**
   * <p>
   * Getter method for the COM property "Resize"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnSize is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getResize(rowSize, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowSize Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(256)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range getResize(
    @Optional java.lang.Object rowSize);

  /**
   * <p>
   * Getter method for the COM property "Resize"
   * </p>
   * @param rowSize Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnSize Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(256)
  @PropGet
  com.exceljava.com4j.excel.Range getResize(
    @Optional java.lang.Object rowSize,
    @Optional java.lang.Object columnSize);


  /**
   * <p>
   * Getter method for the COM property "Row"
   * </p>
   */

  @DISPID(257)
  @PropGet
  int getRow();


  /**
   * @param comparison Mandatory java.lang.Object parameter.
   */

  @DISPID(511)
  com.exceljava.com4j.excel.Range rowDifferences(
    java.lang.Object comparison);


  /**
   * <p>
   * Getter method for the COM property "RowHeight"
   * </p>
   */

  @DISPID(272)
  @PropGet
  java.lang.Object getRowHeight();


  /**
   * <p>
   * Setter method for the COM property "RowHeight"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(272)
  @PropPut
  void setRowHeight(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Rows"
   * </p>
   */

  @DISPID(258)
  @PropGet
  com.exceljava.com4j.excel.Range getRows();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, optParamIndex = {15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, optParamIndex = {16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, optParamIndex = {17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17}, optParamIndex = {18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, optParamIndex = {19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19}, optParamIndex = {20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20}, optParamIndex = {21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21}, optParamIndex = {22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22}, optParamIndex = {23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23}, optParamIndex = {24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23,
    @Optional java.lang.Object arg24);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24}, optParamIndex = {25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23,
    @Optional java.lang.Object arg24,
    @Optional java.lang.Object arg25);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25}, optParamIndex = {26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23,
    @Optional java.lang.Object arg24,
    @Optional java.lang.Object arg25,
    @Optional java.lang.Object arg26);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26}, optParamIndex = {27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23,
    @Optional java.lang.Object arg24,
    @Optional java.lang.Object arg25,
    @Optional java.lang.Object arg26,
    @Optional java.lang.Object arg27);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg28 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27}, optParamIndex = {28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23,
    @Optional java.lang.Object arg24,
    @Optional java.lang.Object arg25,
    @Optional java.lang.Object arg26,
    @Optional java.lang.Object arg27,
    @Optional java.lang.Object arg28);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * run(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg28 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg29 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28}, optParamIndex = {29}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23,
    @Optional java.lang.Object arg24,
    @Optional java.lang.Object arg25,
    @Optional java.lang.Object arg26,
    @Optional java.lang.Object arg27,
    @Optional java.lang.Object arg28,
    @Optional java.lang.Object arg29);

  /**
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg28 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg29 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg30 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(259)
  java.lang.Object run(
    @Optional java.lang.Object arg1,
    @Optional java.lang.Object arg2,
    @Optional java.lang.Object arg3,
    @Optional java.lang.Object arg4,
    @Optional java.lang.Object arg5,
    @Optional java.lang.Object arg6,
    @Optional java.lang.Object arg7,
    @Optional java.lang.Object arg8,
    @Optional java.lang.Object arg9,
    @Optional java.lang.Object arg10,
    @Optional java.lang.Object arg11,
    @Optional java.lang.Object arg12,
    @Optional java.lang.Object arg13,
    @Optional java.lang.Object arg14,
    @Optional java.lang.Object arg15,
    @Optional java.lang.Object arg16,
    @Optional java.lang.Object arg17,
    @Optional java.lang.Object arg18,
    @Optional java.lang.Object arg19,
    @Optional java.lang.Object arg20,
    @Optional java.lang.Object arg21,
    @Optional java.lang.Object arg22,
    @Optional java.lang.Object arg23,
    @Optional java.lang.Object arg24,
    @Optional java.lang.Object arg25,
    @Optional java.lang.Object arg26,
    @Optional java.lang.Object arg27,
    @Optional java.lang.Object arg28,
    @Optional java.lang.Object arg29,
    @Optional java.lang.Object arg30);


  /**
   */

  @DISPID(235)
  java.lang.Object select();


  /**
   */

  @DISPID(496)
  java.lang.Object show();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter remove is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showDependents(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(877)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object showDependents();

  /**
   * @param remove Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(877)
  java.lang.Object showDependents(
    @Optional java.lang.Object remove);


  /**
   * <p>
   * Getter method for the COM property "ShowDetail"
   * </p>
   */

  @DISPID(585)
  @PropGet
  java.lang.Object getShowDetail();


  /**
   * <p>
   * Setter method for the COM property "ShowDetail"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(585)
  @PropPut
  void setShowDetail(
    java.lang.Object rhs);


  /**
   */

  @DISPID(878)
  java.lang.Object showErrors();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter remove is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showPrecedents(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(879)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object showPrecedents();

  /**
   * @param remove Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(879)
  java.lang.Object showPrecedents(
    @Optional java.lang.Object remove);


  /**
   * <p>
   * Getter method for the COM property "ShrinkToFit"
   * </p>
   */

  @DISPID(209)
  @PropGet
  java.lang.Object getShrinkToFit();


  /**
   * <p>
   * Setter method for the COM property "ShrinkToFit"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(209)
  @PropPut
  void setShrinkToFit(
    java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order1 is set to 1</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order1 is set to 1</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, header, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, com4j.Variant.getMissing(), 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, 2, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, 1, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, 0, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, 0, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   * @param dataOption1 Optional parameter. Default value is 0
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2, 0);
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   * @param dataOption1 Optional parameter. Default value is 0
   * @param dataOption2 Optional parameter. Default value is 0
   */

  @DISPID(880)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(index=-1)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2);

  /**
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   * @param dataOption1 Optional parameter. Default value is 0
   * @param dataOption2 Optional parameter. Default value is 0
   * @param dataOption3 Optional parameter. Default value is 0
   */

  @DISPID(880)
  java.lang.Object _Sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption3);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>java.lang.Object parameter key1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order1 is set to 1</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(1, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortMethod.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order1 is set to 1</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order1 is set to 1</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, order3, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, order3, header, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, com4j.Variant.getMissing(), 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, 2, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, 0, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, 0, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param dataOption1 Optional parameter. Default value is 0
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sortSpecial(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2, 0);
   * </code>
   * </p>
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param dataOption1 Optional parameter. Default value is 0
   * @param dataOption2 Optional parameter. Default value is 0
   */

  @DISPID(881)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(index=-1)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2);

  /**
   * @param sortMethod Optional parameter. Default value is 1
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param dataOption1 Optional parameter. Default value is 0
   * @param dataOption2 Optional parameter. Default value is 0
   * @param dataOption3 Optional parameter. Default value is 0
   */

  @DISPID(881)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object type,
    @Optional java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption3);


  /**
   * <p>
   * Getter method for the COM property "SoundNote"
   * </p>
   */

  @DISPID(916)
  @PropGet
  com.exceljava.com4j.excel.SoundNote getSoundNote();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter value is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * specialCells(type, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlCellType parameter.
   */

  @DISPID(410)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Range specialCells(
    com.exceljava.com4j.excel.XlCellType type);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlCellType parameter.
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(410)
  com.exceljava.com4j.excel.Range specialCells(
    com.exceljava.com4j.excel.XlCellType type,
    @Optional java.lang.Object value);


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   */

  @DISPID(260)
  @PropGet
  java.lang.Object getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(260)
  @PropPut
  void setStyle(
    java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSubscribeToFormat parameter format is set to -4158</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * subscribeTo(edition, -4158);
   * </code>
   * </p>
   * @param edition Mandatory java.lang.String parameter.
   */

  @DISPID(481)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlSubscribeToFormat.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-4158"})
  @ReturnValue(index=-1)
  java.lang.Object subscribeTo(
    java.lang.String edition);

  /**
   * @param edition Mandatory java.lang.String parameter.
   * @param format Optional parameter. Default value is -4158
   */

  @DISPID(481)
  java.lang.Object subscribeTo(
    java.lang.String edition,
    @Optional @DefaultValue("-4158") com.exceljava.com4j.excel.XlSubscribeToFormat format);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageBreaks is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSummaryRow parameter summaryBelowData is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * subtotal(groupBy, function, totalList, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1);
   * </code>
   * </p>
   * @param groupBy Mandatory int parameter.
   * @param function Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   * @param totalList Mandatory java.lang.Object parameter.
   */

  @DISPID(882)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSummaryRow.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1"})
  @ReturnValue(index=-1)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    java.lang.Object totalList);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pageBreaks is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSummaryRow parameter summaryBelowData is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * subtotal(groupBy, function, totalList, replace, com4j.Variant.getMissing(), 1);
   * </code>
   * </p>
   * @param groupBy Mandatory int parameter.
   * @param function Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   * @param totalList Mandatory java.lang.Object parameter.
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(882)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSummaryRow.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1"})
  @ReturnValue(index=-1)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    java.lang.Object totalList,
    @Optional java.lang.Object replace);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSummaryRow parameter summaryBelowData is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * subtotal(groupBy, function, totalList, replace, pageBreaks, 1);
   * </code>
   * </p>
   * @param groupBy Mandatory int parameter.
   * @param function Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   * @param totalList Mandatory java.lang.Object parameter.
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageBreaks Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(882)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {com.exceljava.com4j.excel.XlSummaryRow.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=-1)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    java.lang.Object totalList,
    @Optional java.lang.Object replace,
    @Optional java.lang.Object pageBreaks);

  /**
   * @param groupBy Mandatory int parameter.
   * @param function Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   * @param totalList Mandatory java.lang.Object parameter.
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageBreaks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param summaryBelowData Optional parameter. Default value is 1
   */

  @DISPID(882)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    java.lang.Object totalList,
    @Optional java.lang.Object replace,
    @Optional java.lang.Object pageBreaks,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSummaryRow summaryBelowData);


  /**
   * <p>
   * Getter method for the COM property "Summary"
   * </p>
   */

  @DISPID(273)
  @PropGet
  java.lang.Object getSummary();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowInput is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnInput is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * table(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(497)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object table();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnInput is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * table(rowInput, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowInput Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(497)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object table(
    @Optional java.lang.Object rowInput);

  /**
   * @param rowInput Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnInput Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(497)
  java.lang.Object table(
    @Optional java.lang.Object rowInput,
    @Optional java.lang.Object columnInput);


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   */

  @DISPID(138)
  @PropGet
  java.lang.Object getText();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextParsingType parameter dataType is set to 1</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(com4j.Variant.getMissing(), 1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlTextParsingType.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlTextParsingType parameter dataType is set to 1</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, 1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {com.exceljava.com4j.excel.XlTextParsingType.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma,
    @Optional java.lang.Object space);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma,
    @Optional java.lang.Object space,
    @Optional java.lang.Object other);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma,
    @Optional java.lang.Object space,
    @Optional java.lang.Object other,
    @Optional java.lang.Object otherChar);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma,
    @Optional java.lang.Object space,
    @Optional java.lang.Object other,
    @Optional java.lang.Object otherChar,
    @Optional java.lang.Object fieldInfo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma,
    @Optional java.lang.Object space,
    @Optional java.lang.Object other,
    @Optional java.lang.Object otherChar,
    @Optional java.lang.Object fieldInfo,
    @Optional java.lang.Object decimalSeparator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textToColumns(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, thousandsSeparator, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma,
    @Optional java.lang.Object space,
    @Optional java.lang.Object other,
    @Optional java.lang.Object otherChar,
    @Optional java.lang.Object fieldInfo,
    @Optional java.lang.Object decimalSeparator,
    @Optional java.lang.Object thousandsSeparator);

  /**
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is 1
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param trailingMinusNumbers Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1040)
  java.lang.Object textToColumns(
    @Optional java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional java.lang.Object consecutiveDelimiter,
    @Optional java.lang.Object tab,
    @Optional java.lang.Object semicolon,
    @Optional java.lang.Object comma,
    @Optional java.lang.Object space,
    @Optional java.lang.Object other,
    @Optional java.lang.Object otherChar,
    @Optional java.lang.Object fieldInfo,
    @Optional java.lang.Object decimalSeparator,
    @Optional java.lang.Object thousandsSeparator,
    @Optional java.lang.Object trailingMinusNumbers);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   */

  @DISPID(126)
  @PropGet
  java.lang.Object getTop();


  /**
   */

  @DISPID(244)
  java.lang.Object ungroup();


  /**
   * <p>
   * Getter method for the COM property "UseStandardHeight"
   * </p>
   */

  @DISPID(274)
  @PropGet
  java.lang.Object getUseStandardHeight();


  /**
   * <p>
   * Setter method for the COM property "UseStandardHeight"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(274)
  @PropPut
  void setUseStandardHeight(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "UseStandardWidth"
   * </p>
   */

  @DISPID(275)
  @PropGet
  java.lang.Object getUseStandardWidth();


  /**
   * <p>
   * Setter method for the COM property "UseStandardWidth"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(275)
  @PropPut
  void setUseStandardWidth(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Validation"
   * </p>
   */

  @DISPID(1387)
  @PropGet
  com.exceljava.com4j.excel.Validation getValidation();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rangeValueDataType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getValue(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(6)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getValue();

  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @param rangeValueDataType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(6)
  @PropGet
  java.lang.Object getValue(
    @Optional java.lang.Object rangeValueDataType);


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rangeValueDataType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setValue(com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(6)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setValue(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rangeValueDataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(6)
  @PropPut
  void setValue(
    @Optional java.lang.Object rangeValueDataType,
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Value2"
   * </p>
   */

  @DISPID(1388)
  @PropGet
  java.lang.Object getValue2();


  /**
   * <p>
   * Setter method for the COM property "Value2"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1388)
  @PropPut
  void setValue2(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "VerticalAlignment"
   * </p>
   */

  @DISPID(137)
  @PropGet
  java.lang.Object getVerticalAlignment();


  /**
   * <p>
   * Setter method for the COM property "VerticalAlignment"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(137)
  @PropPut
  void setVerticalAlignment(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   */

  @DISPID(122)
  @PropGet
  java.lang.Object getWidth();


  /**
   * <p>
   * Getter method for the COM property "Worksheet"
   * </p>
   */

  @DISPID(348)
  @PropGet
  com.exceljava.com4j.excel._Worksheet getWorksheet();


  /**
   * <p>
   * Getter method for the COM property "WrapText"
   * </p>
   */

  @DISPID(276)
  @PropGet
  java.lang.Object getWrapText();


  /**
   * <p>
   * Setter method for the COM property "WrapText"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(276)
  @PropPut
  void setWrapText(
    java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter text is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addComment(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1389)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Comment addComment();

  /**
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1389)
  com.exceljava.com4j.excel.Comment addComment(
    @Optional java.lang.Object text);


  /**
   * <p>
   * Getter method for the COM property "Comment"
   * </p>
   */

  @DISPID(910)
  @PropGet
  com.exceljava.com4j.excel.Comment getComment();


  /**
   */

  @DISPID(1390)
  void clearComments();


  /**
   * <p>
   * Getter method for the COM property "Phonetic"
   * </p>
   */

  @DISPID(1391)
  @PropGet
  com.exceljava.com4j.excel.Phonetic getPhonetic();


  /**
   * <p>
   * Getter method for the COM property "FormatConditions"
   * </p>
   */

  @DISPID(1392)
  @PropGet
  com.exceljava.com4j.excel.FormatConditions getFormatConditions();


  /**
   * <p>
   * Getter method for the COM property "ReadingOrder"
   * </p>
   */

  @DISPID(975)
  @PropGet
  int getReadingOrder();


  /**
   * <p>
   * Setter method for the COM property "ReadingOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(975)
  @PropPut
  void setReadingOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Hyperlinks"
   * </p>
   */

  @DISPID(1393)
  @PropGet
  com.exceljava.com4j.excel.Hyperlinks getHyperlinks();


  /**
   * <p>
   * Getter method for the COM property "Phonetics"
   * </p>
   */

  @DISPID(1811)
  @PropGet
  com.exceljava.com4j.excel.Phonetics getPhonetics();


  /**
   */

  @DISPID(1812)
  void setPhonetic();


  /**
   * <p>
   * Getter method for the COM property "ID"
   * </p>
   */

  @DISPID(1813)
  @PropGet
  java.lang.String getID();


  /**
   * <p>
   * Setter method for the COM property "ID"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1813)
  @PropPut
  void setID(
    java.lang.String rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing());
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

  @DISPID(1772)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile,
    @Optional java.lang.Object collate);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772)
  java.lang.Object _PrintOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile,
    @Optional java.lang.Object collate,
    @Optional java.lang.Object prToFileName);


  /**
   * <p>
   * Getter method for the COM property "PivotCell"
   * </p>
   */

  @DISPID(2013)
  @PropGet
  com.exceljava.com4j.excel.PivotCell getPivotCell();


  /**
   */

  @DISPID(2014)
  void dirty();


  /**
   * <p>
   * Getter method for the COM property "Errors"
   * </p>
   */

  @DISPID(2015)
  @PropGet
  com.exceljava.com4j.excel.Errors getErrors();


  /**
   * <p>
   * Getter method for the COM property "SmartTags"
   * </p>
   */

  @DISPID(2016)
  @PropGet
  com.exceljava.com4j.excel.SmartTags getSmartTags();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter speakDirection is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter speakFormulas is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * speak(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2017)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void speak();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter speakFormulas is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * speak(speakDirection, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param speakDirection Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2017)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void speak(
    @Optional java.lang.Object speakDirection);

  /**
   * @param speakDirection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param speakFormulas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2017)
  void speak(
    @Optional java.lang.Object speakDirection,
    @Optional java.lang.Object speakFormulas);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPasteType parameter paste is set to -4104</li><li>com.exceljava.com4j.excel.XlPasteSpecialOperation parameter operation is set to -4142</li><li>java.lang.Object parameter skipBlanks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(-4104, -4142, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1928)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteType.class, com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4104", "-4142", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object pasteSpecial();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPasteSpecialOperation parameter operation is set to -4142</li><li>java.lang.Object parameter skipBlanks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(paste, -4142, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param paste Optional parameter. Default value is -4104
   */

  @DISPID(1928)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4142", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object pasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter skipBlanks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(paste, operation, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   */

  @DISPID(1928)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object pasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter transpose is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(paste, operation, skipBlanks, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   * @param skipBlanks Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object pasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional java.lang.Object skipBlanks);

  /**
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   * @param skipBlanks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param transpose Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928)
  java.lang.Object pasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional java.lang.Object skipBlanks,
    @Optional java.lang.Object transpose);


  /**
   * <p>
   * Getter method for the COM property "AllowEdit"
   * </p>
   */

  @DISPID(2020)
  @PropGet
  boolean getAllowEdit();


  /**
   * <p>
   * Getter method for the COM property "ListObject"
   * </p>
   */

  @DISPID(2257)
  @PropGet
  com.exceljava.com4j.excel.ListObject getListObject();


  /**
   * <p>
   * Getter method for the COM property "XPath"
   * </p>
   */

  @DISPID(2258)
  @PropGet
  com.exceljava.com4j.excel.XPath getXPath();


  /**
   * <p>
   * Getter method for the COM property "ServerActions"
   * </p>
   */

  @DISPID(2491)
  @PropGet
  com.exceljava.com4j.excel.Actions getServerActions();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columns is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * removeDuplicates(com4j.Variant.getMissing(), 2);
   * </code>
   * </p>
   */

  @DISPID(2492)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "2"})
  void removeDuplicates();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * removeDuplicates(columns, 2);
   * </code>
   * </p>
   * @param columns Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2492)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"2"})
  void removeDuplicates(
    @Optional java.lang.Object columns);

  /**
   * @param columns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param header Optional parameter. Default value is 2
   */

  @DISPID(2492)
  void removeDuplicates(
    @Optional java.lang.Object columns,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut(
    @Optional java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing());
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

  @DISPID(2361)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object printOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile,
    @Optional java.lang.Object collate);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361)
  java.lang.Object printOut(
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object copies,
    @Optional java.lang.Object preview,
    @Optional java.lang.Object activePrinter,
    @Optional java.lang.Object printToFile,
    @Optional java.lang.Object collate,
    @Optional java.lang.Object prToFileName);


  /**
   * <p>
   * Getter method for the COM property "MDX"
   * </p>
   */

  @DISPID(2123)
  @PropGet
  java.lang.String getMDX();


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

  @DISPID(2493)
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

  @DISPID(2493)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename);

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

  @DISPID(2493)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality);

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

  @DISPID(2493)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties);

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

  @DISPID(2493)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas);

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

  @DISPID(2493)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from);

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

  @DISPID(2493)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from,
    @Optional java.lang.Object to);

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

  @DISPID(2493)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object openAfterPublish);

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

  @DISPID(2493)
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object openAfterPublish,
    @Optional java.lang.Object fixedFormatExtClassPtr);


  /**
   * <p>
   * Getter method for the COM property "CountLarge"
   * </p>
   */

  @DISPID(2499)
  @PropGet
  java.lang.Object getCountLarge();


  /**
   */

  @DISPID(2364)
  java.lang.Object calculateRowMajorOrder();


  /**
   * <p>
   * Getter method for the COM property "SparklineGroups"
   * </p>
   */

  @DISPID(2853)
  @PropGet
  com.exceljava.com4j.excel.SparklineGroups getSparklineGroups();


  /**
   */

  @DISPID(2854)
  void clearHyperlinks();


  /**
   * <p>
   * Getter method for the COM property "DisplayFormat"
   * </p>
   */

  @DISPID(666)
  @PropGet
  com.exceljava.com4j.excel.DisplayFormat getDisplayFormat();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lineStyle is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlBorderWeight parameter weight is set to 2</li><li>com.exceljava.com4j.excel.XlColorIndex parameter colorIndex is set to -4105</li><li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter themeColor is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * borderAround(com4j.Variant.getMissing(), 2, -4105, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2771)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "2", "-4105", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object borderAround();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlBorderWeight parameter weight is set to 2</li><li>com.exceljava.com4j.excel.XlColorIndex parameter colorIndex is set to -4105</li><li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter themeColor is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * borderAround(lineStyle, 2, -4105, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2771)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "-4105", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object borderAround(
    @Optional java.lang.Object lineStyle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlColorIndex parameter colorIndex is set to -4105</li><li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter themeColor is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * borderAround(lineStyle, weight, -4105, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   */

  @DISPID(2771)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4105", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object borderAround(
    @Optional java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter color is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter themeColor is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * borderAround(lineStyle, weight, colorIndex, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   * @param colorIndex Optional parameter. Default value is -4105
   */

  @DISPID(2771)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object borderAround(
    @Optional java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter themeColor is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * borderAround(lineStyle, weight, colorIndex, color, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   * @param colorIndex Optional parameter. Default value is -4105
   * @param color Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2771)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object borderAround(
    @Optional java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex,
    @Optional java.lang.Object color);

  /**
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   * @param colorIndex Optional parameter. Default value is -4105
   * @param color Optional parameter. Default value is com4j.Variant.getMissing()
   * @param themeColor Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2771)
  java.lang.Object borderAround(
    @Optional java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex,
    @Optional java.lang.Object color,
    @Optional java.lang.Object themeColor);


  /**
   */

  @DISPID(2855)
  void allocateChanges();


  /**
   */

  @DISPID(2856)
  void discardChanges();


  /**
   */

  @DISPID(2996)
  void flashFill();


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

  @DISPID(3175)
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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename);

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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality);

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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties);

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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas);

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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from);

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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from,
    @Optional java.lang.Object to);

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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object openAfterPublish);

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

  @DISPID(3175)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object openAfterPublish,
    @Optional java.lang.Object fixedFormatExtClassPtr);

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

  @DISPID(3175)
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional java.lang.Object filename,
    @Optional java.lang.Object quality,
    @Optional java.lang.Object includeDocProperties,
    @Optional java.lang.Object ignorePrintAreas,
    @Optional java.lang.Object from,
    @Optional java.lang.Object to,
    @Optional java.lang.Object openAfterPublish,
    @Optional java.lang.Object fixedFormatExtClassPtr,
    @Optional java.lang.Object workIdentity);


  /**
   * <p>
   * Getter method for the COM property "HasRichDataType"
   * </p>
   */

  @DISPID(3320)
  @PropGet
  java.lang.Object getHasRichDataType();


  /**
   */

  @DISPID(3274)
  void showCard();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter text is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCommentThreaded(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(3280)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CommentThreaded addCommentThreaded();

  /**
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3280)
  com.exceljava.com4j.excel.CommentThreaded addCommentThreaded(
    @Optional java.lang.Object text);


  /**
   * <p>
   * Getter method for the COM property "CommentThreaded"
   * </p>
   */

  @DISPID(3281)
  @PropGet
  com.exceljava.com4j.excel.CommentThreaded getCommentThreaded();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order1 is set to 1</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order1 is set to 1</li><li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order2 is set to 1</li><li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, 1, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter key3 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, com4j.Variant.getMissing(), 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrder parameter order3 is set to 1</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, 1, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlYesNoGuess parameter header is set to 2</li><li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter orderCustom is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, com4j.Variant.getMissing(), 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortOrientation parameter orientation is set to 2</li><li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, 2, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortMethod parameter sortMethod is set to 1</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, 1, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption1 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, 0, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption2 is set to 0</li><li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, 0, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   * @param dataOption1 Optional parameter. Default value is 0
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSortDataOption parameter dataOption3 is set to 0</li><li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   * @param dataOption1 Optional parameter. Default value is 0
   * @param dataOption2 Optional parameter. Default value is 0
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14, 15}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subField1 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sort(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2, dataOption3, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   * @param dataOption1 Optional parameter. Default value is 0
   * @param dataOption2 Optional parameter. Default value is 0
   * @param dataOption3 Optional parameter. Default value is 0
   */

  @DISPID(3288)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, optParamIndex = {15}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption3);

  /**
   * @param key1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order1 Optional parameter. Default value is 1
   * @param key2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order2 Optional parameter. Default value is 1
   * @param key3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order3 Optional parameter. Default value is 1
   * @param header Optional parameter. Default value is 2
   * @param orderCustom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param orientation Optional parameter. Default value is 2
   * @param sortMethod Optional parameter. Default value is 1
   * @param dataOption1 Optional parameter. Default value is 0
   * @param dataOption2 Optional parameter. Default value is 0
   * @param dataOption3 Optional parameter. Default value is 0
   * @param subField1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3288)
  java.lang.Object sort(
    @Optional java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional java.lang.Object key2,
    @Optional java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional java.lang.Object orderCustom,
    @Optional java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption3,
    @Optional java.lang.Object subField1);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter field is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter criteria1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlAutoFilterOperator parameter operator is set to 1</li><li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFilter(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(3289)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFilter();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter criteria1 is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlAutoFilterOperator parameter operator is set to 1</li><li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFilter(field, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3289)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFilter(
    @Optional java.lang.Object field);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlAutoFilterOperator parameter operator is set to 1</li><li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFilter(field, criteria1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3289)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter criteria2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFilter(field, criteria1, operator, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   */

  @DISPID(3289)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter visibleDropDown is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFilter(field, criteria1, operator, criteria2, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   * @param criteria2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3289)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional java.lang.Object criteria2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFilter(field, criteria1, operator, criteria2, visibleDropDown, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   * @param criteria2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visibleDropDown Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3289)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object autoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional java.lang.Object criteria2,
    @Optional java.lang.Object visibleDropDown);

  /**
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   * @param criteria2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visibleDropDown Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subField Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3289)
  java.lang.Object autoFilter(
    @Optional java.lang.Object field,
    @Optional java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional java.lang.Object criteria2,
    @Optional java.lang.Object visibleDropDown,
    @Optional java.lang.Object subField);


  /**
   * @param serviceID Mandatory int parameter.
   * @param languageCulture Mandatory java.lang.String parameter.
   */

  @DISPID(3290)
  void convertToLinkedDataType(
    int serviceID,
    java.lang.String languageCulture);


  /**
   * <p>
   * Getter method for the COM property "LinkedDataTypeState"
   * </p>
   */

  @DISPID(3291)
  @PropGet
  java.lang.Object getLinkedDataTypeState();


  /**
   * @param sourceCell Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(3293)
  void setCellDataTypeFromCell(
    com.exceljava.com4j.excel.Range sourceCell);


  /**
   */

  @DISPID(3294)
  void dataTypeToText();


  /**
   * <p>
   * Getter method for the COM property "HasSpill"
   * </p>
   */

  @DISPID(3295)
  @PropGet
  java.lang.Object getHasSpill();


  /**
   * <p>
   * Getter method for the COM property "SpillingToRange"
   * </p>
   */

  @DISPID(3296)
  @PropGet
  com.exceljava.com4j.excel.Range getSpillingToRange();


  /**
   * <p>
   * Getter method for the COM property "SpillParent"
   * </p>
   */

  @DISPID(3297)
  @PropGet
  com.exceljava.com4j.excel.Range getSpillParent();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter domainID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * refreshLinkedDataType(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(3299)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void refreshLinkedDataType();

  /**
   * @param domainID Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3299)
  void refreshLinkedDataType(
    @Optional java.lang.Object domainID);


  /**
   * <p>
   * Getter method for the COM property "Formula2"
   * </p>
   */

  @DISPID(1580)
  @PropGet
  java.lang.Object getFormula2();


  /**
   * <p>
   * Setter method for the COM property "Formula2"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1580)
  @PropPut
  void setFormula2(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula2Local"
   * </p>
   */

  @DISPID(3300)
  @PropGet
  java.lang.Object getFormula2Local();


  /**
   * <p>
   * Setter method for the COM property "Formula2Local"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(3300)
  @PropPut
  void setFormula2Local(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula2R1C1"
   * </p>
   */

  @DISPID(3301)
  @PropGet
  java.lang.Object getFormula2R1C1();


  /**
   * <p>
   * Setter method for the COM property "Formula2R1C1"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(3301)
  @PropPut
  void setFormula2R1C1(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula2R1C1Local"
   * </p>
   */

  @DISPID(3302)
  @PropGet
  java.lang.Object getFormula2R1C1Local();


  /**
   * <p>
   * Setter method for the COM property "Formula2R1C1Local"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(3302)
  @PropPut
  void setFormula2R1C1Local(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SavedAsArray"
   * </p>
   */

  @DISPID(3303)
  @PropGet
  java.lang.Object getSavedAsArray();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lookAt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formulaVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(what, replacement, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   */

  @DISPID(3305)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter searchOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formulaVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(what, replacement, lookAt, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3305)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchCase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formulaVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(what, replacement, lookAt, searchOrder, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3305)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter matchByte is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formulaVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(what, replacement, lookAt, searchOrder, matchCase, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3305)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter searchFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formulaVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(what, replacement, lookAt, searchOrder, matchCase, matchByte, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3305)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replaceFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formulaVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3305)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte,
    @Optional java.lang.Object searchFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formulaVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat, replaceFormat, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replaceFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3305)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte,
    @Optional java.lang.Object searchFormat,
    @Optional java.lang.Object replaceFormat);

  /**
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replaceFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formulaVersion Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3305)
  boolean replace(
    java.lang.Object what,
    java.lang.Object replacement,
    @Optional java.lang.Object lookAt,
    @Optional java.lang.Object searchOrder,
    @Optional java.lang.Object matchCase,
    @Optional java.lang.Object matchByte,
    @Optional java.lang.Object searchFormat,
    @Optional java.lang.Object replaceFormat,
    @Optional java.lang.Object formulaVersion);


  // Properties:
}
