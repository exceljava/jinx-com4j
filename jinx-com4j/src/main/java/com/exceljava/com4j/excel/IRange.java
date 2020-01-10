package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020846-0001-0000-C000-000000000046}")
public interface IRange extends Com4jObject,Iterable<Com4jObject> {
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object activate();


  /**
   * <p>
   * Getter method for the COM property "AddIndent"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAddIndent();


  /**
   * <p>
   * Setter method for the COM property "AddIndent"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(12)
  void setAddIndent(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowAbsolute is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnAbsolute is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {6}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "1033"})
  @ReturnValue(index=6)
  java.lang.String getAddress();

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnAbsolute is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 6}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "1033"})
  @ReturnValue(index=6)
  java.lang.String getAddress(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlReferenceStyle parameter referenceStyle is set to 1</li><li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, columnAbsolute, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "1033"})
  @ReturnValue(index=6)
  java.lang.String getAddress(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter external is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, columnAbsolute, referenceStyle, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(index=6)
  java.lang.String getAddress(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter relativeTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, columnAbsolute, referenceStyle, external, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=6)
  java.lang.String getAddress(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object external);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAddress(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo, 1033);
   * </code>
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   * @param relativeTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=6)
  java.lang.String getAddress(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object external,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object relativeTo);

  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   * @param relativeTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getAddress(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object external,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object relativeTo,
    @LCID int lcid);


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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=5)
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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004"})
  @ReturnValue(index=5)
  java.lang.String getAddressLocal(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute);

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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlReferenceStyle.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004"})
  @ReturnValue(index=5)
  java.lang.String getAddressLocal(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute);

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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  java.lang.String getAddressLocal(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute,
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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  java.lang.String getAddressLocal(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object external);

  /**
   * <p>
   * Getter method for the COM property "AddressLocal"
   * </p>
   * @param rowAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param referenceStyle Optional parameter. Default value is 1
   * @param external Optional parameter. Default value is com4j.Variant.getMissing()
   * @param relativeTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  java.lang.String getAddressLocal(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnAbsolute,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlReferenceStyle referenceStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object external,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object relativeTo);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object advancedFilter(
    com.exceljava.com4j.excel.XlFilterAction action,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteriaRange);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object advancedFilter(
    com.exceljava.com4j.excel.XlFilterAction action,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteriaRange,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copyToRange);

  /**
   * @param action Mandatory com.exceljava.com4j.excel.XlFilterAction parameter.
   * @param criteriaRange Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copyToRange Optional parameter. Default value is com4j.Variant.getMissing()
   * @param unique Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object advancedFilter(
    com.exceljava.com4j.excel.XlFilterAction action,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteriaRange,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copyToRange,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object unique);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {7}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 7}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object applyNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object names);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object applyNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object names,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreRelativeAbsolute);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object applyNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object names,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreRelativeAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useRowColumnNames);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object applyNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object names,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreRelativeAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useRowColumnNames,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object omitColumn);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {com.exceljava.com4j.excel.XlApplyNamesOrder.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object applyNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object names,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreRelativeAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useRowColumnNames,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object omitColumn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object omitRow);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object applyNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object names,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreRelativeAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useRowColumnNames,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object omitColumn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object omitRow,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlApplyNamesOrder order);

  /**
   * @param names Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreRelativeAbsolute Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useRowColumnNames Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitColumn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param omitRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is 1
   * @param appendLast Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object applyNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object names,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreRelativeAbsolute,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useRowColumnNames,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object omitColumn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object omitRow,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlApplyNamesOrder order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object appendLast);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(17)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object applyOutlineStyles();


  /**
   * <p>
   * Getter method for the COM property "Areas"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Areas
   */

  @VTID(18)
  com.exceljava.com4j.excel.Areas getAreas();


  /**
   * @param string Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.String
   */

  @VTID(19)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlAutoFillType.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object autoFill(
    com.exceljava.com4j.excel.Range destination);

  /**
   * @param destination Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param type Optional parameter. Default value is 0
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(20)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object _AutoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object _AutoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object _AutoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object _AutoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria2);

  /**
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   * @param criteria2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visibleDropDown Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _AutoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visibleDropDown);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {7}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {com.exceljava.com4j.excel.XlRangeAutoFormat.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 7}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object number);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object number,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object font);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object number,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object font,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alignment);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object number,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object font,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alignment,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object border);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object number,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object font,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alignment,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object border,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pattern);

  /**
   * @param format Optional parameter. Default value is 1
   * @param number Optional parameter. Default value is com4j.Variant.getMissing()
   * @param font Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alignment Optional parameter. Default value is com4j.Variant.getMissing()
   * @param border Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pattern Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object autoFormat(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlRangeAutoFormat format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object number,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object font,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alignment,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object border,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pattern,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(24)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "2", "-4105", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"2", "-4105", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object _BorderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"-4105", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object _BorderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object _BorderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex);

  /**
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   * @param colorIndex Optional parameter. Default value is -4105
   * @param color Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(25)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _BorderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object color);


  /**
   * <p>
   * Getter method for the COM property "Borders"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Borders
   */

  @VTID(26)
  com.exceljava.com4j.excel.Borders getBorders();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(27)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object calculate();


  /**
   * <p>
   * Getter method for the COM property "Cells"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(28)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start);

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(29)
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object length);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest);

  /**
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(30)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(31)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clear();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(32)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearContents();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(33)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearFormats();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(34)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearNotes();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(35)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearOutline();


  /**
   * <p>
   * Getter method for the COM property "Column"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(36)
  int getColumn();


  /**
   * @param comparison Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(37)
  com.exceljava.com4j.excel.Range columnDifferences(
    @MarshalAs(NativeType.VARIANT) java.lang.Object comparison);


  /**
   * <p>
   * Getter method for the COM property "Columns"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(38)
  com.exceljava.com4j.excel.Range getColumns();


  /**
   * <p>
   * Getter method for the COM property "ColumnWidth"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(39)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "ColumnWidth"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(40)
  void setColumnWidth(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object consolidate(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sources);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object consolidate(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sources,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object function);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object consolidate(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sources,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object function,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object topRow);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object consolidate(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sources,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object function,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object topRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object leftColumn);

  /**
   * @param sources Optional parameter. Default value is com4j.Variant.getMissing()
   * @param function Optional parameter. Default value is com4j.Variant.getMissing()
   * @param topRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param leftColumn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(41)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object consolidate(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sources,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object function,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object topRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object leftColumn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createLinks);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object copy();

  /**
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(42)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object copy(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);


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
   * @return  Returns a value of type int
   */

  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
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
   * @return  Returns a value of type int
   */

  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  int copyFromRecordset(
    com4j.Com4jObject data,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object maxRows);

  /**
   * @param data Mandatory com4j.Com4jObject parameter.
   * @param maxRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param maxColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type int
   */

  @VTID(43)
  int copyFromRecordset(
    com4j.Com4jObject data,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object maxRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object maxColumns);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(44)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "-4147"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(44)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-4147"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * @param appearance Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is -4147
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(44)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(45)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object createNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object createNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object createNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object bottom);

  /**
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param bottom Optional parameter. Default value is com4j.Variant.getMissing()
   * @param right Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(46)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object createNames(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object bottom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object right);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {6}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 6}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsBIFF);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsBIFF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsRTF);

  /**
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsRTF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsVALU Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsBIFF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsRTF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsVALU);


  /**
   * <p>
   * Getter method for the COM property "CurrentArray"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(48)
  com.exceljava.com4j.excel.Range getCurrentArray();


  /**
   * <p>
   * Getter method for the COM property "CurrentRegion"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(49)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(50)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object cut();

  /**
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(50)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object cut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {6}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlDataSeriesType.class, com.exceljava.com4j.excel.XlDataSeriesDate.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "-4132", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {0, 6}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlDataSeriesType.class, com.exceljava.com4j.excel.XlDataSeriesDate.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4132", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object dataSeries(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlDataSeriesDate.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object dataSeries(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object dataSeries(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object dataSeries(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlDataSeriesDate date,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object step);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object dataSeries(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlDataSeriesDate date,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object step,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object stop);

  /**
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is -4132
   * @param date Optional parameter. Default value is 1
   * @param step Optional parameter. Default value is com4j.Variant.getMissing()
   * @param stop Optional parameter. Default value is com4j.Variant.getMissing()
   * @param trend Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(51)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object dataSeries(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol,
    @Optional @DefaultValue("-4132") com.exceljava.com4j.excel.XlDataSeriesType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlDataSeriesDate date,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object step,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object stop,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object trend);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object get_Default();

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(rowIndex, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object get_Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex);

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(rowIndex, columnIndex, 1033);
   * </code>
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object get_Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex);

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @DefaultMethod
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object get_Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex,
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * set_Default(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(53)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void set_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * set_Default(rowIndex, com4j.Variant.getMissing(), 1033, rhs);
   * </code>
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(53)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void set_Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * set_Default(rowIndex, columnIndex, 1033, rhs);
   * </code>
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(53)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void set_Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rowIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(53)
  @DefaultMethod
  void set_Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex,
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(54)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object delete();

  /**
   * @param shift Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(54)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shift);


  /**
   * <p>
   * Getter method for the COM property "Dependents"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(55)
  com.exceljava.com4j.excel.Range getDependents();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object dialogBox();


  /**
   * <p>
   * Getter method for the COM property "DirectDependents"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(57)
  com.exceljava.com4j.excel.Range getDirectDependents();


  /**
   * <p>
   * Getter method for the COM property "DirectPrecedents"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(58)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reference);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reference,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reference,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(59)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object editionOptions(
    com.exceljava.com4j.excel.XlEditionType type,
    com.exceljava.com4j.excel.XlEditionOptionsOption option,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reference,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance chartSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);


  /**
   * <p>
   * Getter method for the COM property "End"
   * </p>
   * @param direction Mandatory com.exceljava.com4j.excel.XlDirection parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(60)
  com.exceljava.com4j.excel.Range getEnd(
    com.exceljava.com4j.excel.XlDirection direction);


  /**
   * <p>
   * Getter method for the COM property "EntireColumn"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(61)
  com.exceljava.com4j.excel.Range getEntireColumn();


  /**
   * <p>
   * Getter method for the COM property "EntireRow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(62)
  com.exceljava.com4j.excel.Range getEntireRow();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(63)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object fillDown();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(64)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object fillLeft();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(65)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object fillRight();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(66)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 9}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 1, 9}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 9}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookIn);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 9}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookIn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 9}, optParamIndex = {5, 6, 7, 8}, javaType = {com.exceljava.com4j.excel.XlSearchDirection.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookIn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 9}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookIn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 9}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookIn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSearchDirection searchDirection,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 9}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookIn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSearchDirection searchDirection,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(67)
  com.exceljava.com4j.excel.Range find(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookIn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSearchDirection searchDirection,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchFormat);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(68)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Range findNext();

  /**
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(68)
  com.exceljava.com4j.excel.Range findNext(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(69)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Range findPrevious();

  /**
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(69)
  com.exceljava.com4j.excel.Range findPrevious(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after);


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Font
   */

  @VTID(70)
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFormula(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(71)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getFormula();

  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(71)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormula(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setFormula(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(72)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setFormula(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(72)
  void setFormula(
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaArray"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(73)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormulaArray();


  /**
   * <p>
   * Setter method for the COM property "FormulaArray"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(74)
  void setFormulaArray(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaLabel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlFormulaLabel
   */

  @VTID(75)
  com.exceljava.com4j.excel.XlFormulaLabel getFormulaLabel();


  /**
   * <p>
   * Setter method for the COM property "FormulaLabel"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlFormulaLabel parameter.
   */

  @VTID(76)
  void setFormulaLabel(
    com.exceljava.com4j.excel.XlFormulaLabel rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaHidden"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(77)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormulaHidden();


  /**
   * <p>
   * Setter method for the COM property "FormulaHidden"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(78)
  void setFormulaHidden(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaLocal"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(79)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormulaLocal();


  /**
   * <p>
   * Setter method for the COM property "FormulaLocal"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(80)
  void setFormulaLocal(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaR1C1"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFormulaR1C1(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(81)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getFormulaR1C1();

  /**
   * <p>
   * Getter method for the COM property "FormulaR1C1"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(81)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormulaR1C1(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "FormulaR1C1"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setFormulaR1C1(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(82)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setFormulaR1C1(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "FormulaR1C1"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(82)
  void setFormulaR1C1(
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaR1C1Local"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(83)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormulaR1C1Local();


  /**
   * <p>
   * Setter method for the COM property "FormulaR1C1Local"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(84)
  void setFormulaR1C1Local(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(85)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object functionWizard();


  /**
   * @param goal Mandatory java.lang.Object parameter.
   * @param changingCell Mandatory com.exceljava.com4j.excel.Range parameter.
   * @return  Returns a value of type boolean
   */

  @VTID(86)
  boolean goalSeek(
    @MarshalAs(NativeType.VARIANT) java.lang.Object goal,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object group(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object group(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object end);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object group(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object end,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object by);

  /**
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param end Optional parameter. Default value is com4j.Variant.getMissing()
   * @param by Optional parameter. Default value is com4j.Variant.getMissing()
   * @param periods Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(87)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object group(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object end,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object by,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object periods);


  /**
   * <p>
   * Getter method for the COM property "HasArray"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(88)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHasArray();


  /**
   * <p>
   * Getter method for the COM property "HasFormula"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(89)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHasFormula();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(90)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHeight();


  /**
   * <p>
   * Getter method for the COM property "Hidden"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(91)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHidden();


  /**
   * <p>
   * Setter method for the COM property "Hidden"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(92)
  void setHidden(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalAlignment"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(93)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHorizontalAlignment();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAlignment"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(94)
  void setHorizontalAlignment(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "IndentLevel"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(95)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getIndentLevel();


  /**
   * <p>
   * Setter method for the COM property "IndentLevel"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(96)
  void setIndentLevel(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * @param insertAmount Mandatory int parameter.
   */

  @VTID(97)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(98)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(98)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object insert(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shift);

  /**
   * @param shift Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copyOrigin Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(98)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object insert(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shift,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copyOrigin);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Interior
   */

  @VTID(99)
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getItem(rowIndex, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(100)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex);

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getItem(rowIndex, columnIndex, 1033);
   * </code>
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(100)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex);

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(100)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex,
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnIndex is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setItem(rowIndex, com4j.Variant.getMissing(), 1033, rhs);
   * </code>
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(101)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void setItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setItem(rowIndex, columnIndex, 1033, rhs);
   * </code>
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(101)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Item"
   * </p>
   * @param rowIndex Mandatory java.lang.Object parameter.
   * @param columnIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(101)
  void setItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rowIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnIndex,
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(102)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object justify();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(103)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getLeft();


  /**
   * <p>
   * Getter method for the COM property "ListHeaderRows"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(104)
  int getListHeaderRows();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(105)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object listNames();


  /**
   * <p>
   * Getter method for the COM property "LocationInTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlLocationInTable
   */

  @VTID(106)
  com.exceljava.com4j.excel.XlLocationInTable getLocationInTable();


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(107)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(108)
  void setLocked(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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

  @VTID(109)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void merge();

  /**
   * @param across Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(109)
  void merge(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object across);


  /**
   */

  @VTID(110)
  void unMerge();


  /**
   * <p>
   * Getter method for the COM property "MergeArea"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(111)
  com.exceljava.com4j.excel.Range getMergeArea();


  /**
   * <p>
   * Getter method for the COM property "MergeCells"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(112)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getMergeCells();


  /**
   * <p>
   * Setter method for the COM property "MergeCells"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(113)
  void setMergeCells(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(114)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(115)
  void setName(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object navigateArrow(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object towardPrecedent);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object navigateArrow(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object towardPrecedent,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arrowNumber);

  /**
   * @param towardPrecedent Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arrowNumber Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkNumber Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(116)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object navigateArrow(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object towardPrecedent,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arrowNumber,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkNumber);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(117)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Next"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(118)
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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(119)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=3)
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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(119)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  java.lang.String noteText(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object text);

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
   * @return  Returns a value of type java.lang.String
   */

  @VTID(119)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  java.lang.String noteText(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object text,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start);

  /**
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @VTID(119)
  java.lang.String noteText(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object text,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object length);


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(120)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(121)
  void setNumberFormat(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormatLocal"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(122)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getNumberFormatLocal();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLocal"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(123)
  void setNumberFormatLocal(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(124)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(124)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Range getOffset(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowOffset);

  /**
   * <p>
   * Getter method for the COM property "Offset"
   * </p>
   * @param rowOffset Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnOffset Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(124)
  com.exceljava.com4j.excel.Range getOffset(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowOffset,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnOffset);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(125)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(126)
  void setOrientation(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "OutlineLevel"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(127)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getOutlineLevel();


  /**
   * <p>
   * Setter method for the COM property "OutlineLevel"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(128)
  void setOutlineLevel(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PageBreak"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(129)
  int getPageBreak();


  /**
   * <p>
   * Setter method for the COM property "PageBreak"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(130)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(131)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(131)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object parse(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parseLine);

  /**
   * @param parseLine Optional parameter. Default value is com4j.Variant.getMissing()
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(131)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object parse(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parseLine,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(132)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteType.class, com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4104", "-4142", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(132)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4142", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(132)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(132)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object _PasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object skipBlanks);

  /**
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   * @param skipBlanks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param transpose Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(132)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _PasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object skipBlanks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object transpose);


  /**
   * <p>
   * Getter method for the COM property "PivotField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(133)
  com.exceljava.com4j.excel.PivotField getPivotField();


  /**
   * <p>
   * Getter method for the COM property "PivotItem"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotItem
   */

  @VTID(134)
  com.exceljava.com4j.excel.PivotItem getPivotItem();


  /**
   * <p>
   * Getter method for the COM property "PivotTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @VTID(135)
  com.exceljava.com4j.excel.PivotTable getPivotTable();


  /**
   * <p>
   * Getter method for the COM property "Precedents"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(136)
  com.exceljava.com4j.excel.Range getPrecedents();


  /**
   * <p>
   * Getter method for the COM property "PrefixCharacter"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(137)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getPrefixCharacter();


  /**
   * <p>
   * Getter method for the COM property "Previous"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(138)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {7}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {0, 7}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object __PrintOut(
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=7)
  java.lang.Object __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(139)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object __PrintOut(
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
     * <li>java.lang.Object parameter enableChanges is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printPreview(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(140)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object printPreview();

  /**
   * @param enableChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(140)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object printPreview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enableChanges);


  /**
   * <p>
   * Getter method for the COM property "QueryTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._QueryTable
   */

  @VTID(141)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(142)
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

  @VTID(142)
  com.exceljava.com4j.excel.Range getRange(
    @MarshalAs(NativeType.VARIANT) java.lang.Object cell1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object cell2);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(143)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type boolean
   */

  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  boolean _Replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement);

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
   * @return  Returns a value of type boolean
   */

  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  boolean _Replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt);

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
   * @return  Returns a value of type boolean
   */

  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  boolean _Replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder);

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
   * @return  Returns a value of type boolean
   */

  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  boolean _Replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase);

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
   * @return  Returns a value of type boolean
   */

  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=8)
  boolean _Replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte);

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
   * @return  Returns a value of type boolean
   */

  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=8)
  boolean _Replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchFormat);

  /**
   * @param what Mandatory java.lang.Object parameter.
   * @param replacement Mandatory java.lang.Object parameter.
   * @param lookAt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchCase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param matchByte Optional parameter. Default value is com4j.Variant.getMissing()
   * @param searchFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replaceFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(144)
  boolean _Replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replaceFormat);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(145)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(145)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Range getResize(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowSize);

  /**
   * <p>
   * Getter method for the COM property "Resize"
   * </p>
   * @param rowSize Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnSize Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(145)
  com.exceljava.com4j.excel.Range getResize(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnSize);


  /**
   * <p>
   * Getter method for the COM property "Row"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(146)
  int getRow();


  /**
   * @param comparison Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(147)
  com.exceljava.com4j.excel.Range rowDifferences(
    @MarshalAs(NativeType.VARIANT) java.lang.Object comparison);


  /**
   * <p>
   * Getter method for the COM property "RowHeight"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(148)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getRowHeight();


  /**
   * <p>
   * Setter method for the COM property "RowHeight"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(149)
  void setRowHeight(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Rows"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(150)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {30}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 30}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 30}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 30}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 30}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 30}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 30}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 30}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 30}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 30}, optParamIndex = {9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 30}, optParamIndex = {10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 30}, optParamIndex = {11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 30}, optParamIndex = {12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 30}, optParamIndex = {13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 30}, optParamIndex = {14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 30}, optParamIndex = {15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 30}, optParamIndex = {16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 30}, optParamIndex = {17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 30}, optParamIndex = {18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 30}, optParamIndex = {19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 30}, optParamIndex = {20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 30}, optParamIndex = {21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 30}, optParamIndex = {22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 30}, optParamIndex = {23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 30}, optParamIndex = {24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 30}, optParamIndex = {25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 30}, optParamIndex = {26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 30}, optParamIndex = {27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 30}, optParamIndex = {28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg28);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 30}, optParamIndex = {29}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=30)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg28,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg29);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(151)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object run(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg28,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg29,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg30);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(152)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(153)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object showDependents();

  /**
   * @param remove Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(154)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object showDependents(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object remove);


  /**
   * <p>
   * Getter method for the COM property "ShowDetail"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(155)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getShowDetail();


  /**
   * <p>
   * Setter method for the COM property "ShowDetail"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(156)
  void setShowDetail(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(157)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(158)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object showPrecedents();

  /**
   * @param remove Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(158)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object showPrecedents(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object remove);


  /**
   * <p>
   * Getter method for the COM property "ShrinkToFit"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(159)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getShrinkToFit();


  /**
   * <p>
   * Setter method for the COM property "ShrinkToFit"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(160)
  void setShrinkToFit(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {15}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 15}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 15}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 15}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 15}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 15}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 15}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 15}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 15}, optParamIndex = {8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 15}, optParamIndex = {9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 15}, optParamIndex = {10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15}, optParamIndex = {11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 15}, optParamIndex = {12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 15}, optParamIndex = {13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15}, optParamIndex = {14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(161)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _Sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {15}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortMethod.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 15}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 15}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 15}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 15}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 15}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 15}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 15}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 15}, optParamIndex = {8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 15}, optParamIndex = {9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 15}, optParamIndex = {10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15}, optParamIndex = {11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 15}, optParamIndex = {12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 15}, optParamIndex = {13, 14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15}, optParamIndex = {14}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(type=NativeType.VARIANT,index=15)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(162)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object sortSpecial(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption3);


  /**
   * <p>
   * Getter method for the COM property "SoundNote"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SoundNote
   */

  @VTID(163)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(164)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Range specialCells(
    com.exceljava.com4j.excel.XlCellType type);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlCellType parameter.
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(164)
  com.exceljava.com4j.excel.Range specialCells(
    com.exceljava.com4j.excel.XlCellType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(165)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(166)
  void setStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(167)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlSubscribeToFormat.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-4158"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object subscribeTo(
    java.lang.String edition);

  /**
   * @param edition Mandatory java.lang.String parameter.
   * @param format Optional parameter. Default value is -4158
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(167)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(168)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSummaryRow.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    @MarshalAs(NativeType.VARIANT) java.lang.Object totalList);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(168)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSummaryRow.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    @MarshalAs(NativeType.VARIANT) java.lang.Object totalList,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(168)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {com.exceljava.com4j.excel.XlSummaryRow.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    @MarshalAs(NativeType.VARIANT) java.lang.Object totalList,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageBreaks);

  /**
   * @param groupBy Mandatory int parameter.
   * @param function Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   * @param totalList Mandatory java.lang.Object parameter.
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageBreaks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param summaryBelowData Optional parameter. Default value is 1
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(168)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object subtotal(
    int groupBy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    @MarshalAs(NativeType.VARIANT) java.lang.Object totalList,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageBreaks,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSummaryRow summaryBelowData);


  /**
   * <p>
   * Getter method for the COM property "Summary"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(169)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(170)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(170)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object table(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowInput);

  /**
   * @param rowInput Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnInput Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(170)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object table(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowInput,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnInput);


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(171)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {14}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlTextParsingType.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 14}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {com.exceljava.com4j.excel.XlTextParsingType.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 14}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 14}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 14}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 14}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 14}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 14}, optParamIndex = {7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 14}, optParamIndex = {8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 14}, optParamIndex = {9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 14}, optParamIndex = {10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 14}, optParamIndex = {11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 14}, optParamIndex = {12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14}, optParamIndex = {13}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=14)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(172)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object textToColumns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextParsingType dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object trailingMinusNumbers);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(173)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getTop();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(174)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object ungroup();


  /**
   * <p>
   * Getter method for the COM property "UseStandardHeight"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(175)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getUseStandardHeight();


  /**
   * <p>
   * Setter method for the COM property "UseStandardHeight"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(176)
  void setUseStandardHeight(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "UseStandardWidth"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(177)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getUseStandardWidth();


  /**
   * <p>
   * Setter method for the COM property "UseStandardWidth"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(178)
  void setUseStandardWidth(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Validation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Validation
   */

  @VTID(179)
  com.exceljava.com4j.excel.Validation getValidation();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rangeValueDataType is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getValue(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(180)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object getValue();

  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getValue(rangeValueDataType, 1033);
   * </code>
   * </p>
   * @param rangeValueDataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(180)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object getValue(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rangeValueDataType);

  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @param rangeValueDataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(180)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rangeValueDataType,
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rangeValueDataType is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setValue(com4j.Variant.getMissing(), 1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(181)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void setValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setValue(rangeValueDataType, 1033, rhs);
   * </code>
   * </p>
   * @param rangeValueDataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(181)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setValue(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rangeValueDataType,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rangeValueDataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(181)
  void setValue(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rangeValueDataType,
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Value2"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getValue2(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(182)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getValue2();

  /**
   * <p>
   * Getter method for the COM property "Value2"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(182)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue2(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Value2"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setValue2(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(183)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setValue2(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Value2"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(183)
  void setValue2(
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "VerticalAlignment"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(184)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVerticalAlignment();


  /**
   * <p>
   * Setter method for the COM property "VerticalAlignment"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(185)
  void setVerticalAlignment(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(186)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getWidth();


  /**
   * <p>
   * Getter method for the COM property "Worksheet"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Worksheet
   */

  @VTID(187)
  com.exceljava.com4j.excel._Worksheet getWorksheet();


  /**
   * <p>
   * Getter method for the COM property "WrapText"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(188)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getWrapText();


  /**
   * <p>
   * Setter method for the COM property "WrapText"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(189)
  void setWrapText(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.Comment
   */

  @VTID(190)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Comment addComment();

  /**
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Comment
   */

  @VTID(190)
  com.exceljava.com4j.excel.Comment addComment(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object text);


  /**
   * <p>
   * Getter method for the COM property "Comment"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Comment
   */

  @VTID(191)
  com.exceljava.com4j.excel.Comment getComment();


  /**
   */

  @VTID(192)
  void clearComments();


  /**
   * <p>
   * Getter method for the COM property "Phonetic"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Phonetic
   */

  @VTID(193)
  com.exceljava.com4j.excel.Phonetic getPhonetic();


  /**
   * <p>
   * Getter method for the COM property "FormatConditions"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.FormatConditions
   */

  @VTID(194)
  com.exceljava.com4j.excel.FormatConditions getFormatConditions();


  /**
   * <p>
   * Getter method for the COM property "ReadingOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(195)
  int getReadingOrder();


  /**
   * <p>
   * Setter method for the COM property "ReadingOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(196)
  void setReadingOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Hyperlinks"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Hyperlinks
   */

  @VTID(197)
  com.exceljava.com4j.excel.Hyperlinks getHyperlinks();


  /**
   * <p>
   * Getter method for the COM property "Phonetics"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Phonetics
   */

  @VTID(198)
  com.exceljava.com4j.excel.Phonetics getPhonetics();


  /**
   */

  @VTID(199)
  void setPhonetic();


  /**
   * <p>
   * Getter method for the COM property "ID"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(200)
  java.lang.String getID();


  /**
   * <p>
   * Setter method for the COM property "ID"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(201)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {8}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
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
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(202)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _PrintOut(
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
   * Getter method for the COM property "PivotCell"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCell
   */

  @VTID(203)
  com.exceljava.com4j.excel.PivotCell getPivotCell();


  /**
   */

  @VTID(204)
  void dirty();


  /**
   * <p>
   * Getter method for the COM property "Errors"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Errors
   */

  @VTID(205)
  com.exceljava.com4j.excel.Errors getErrors();


  /**
   * <p>
   * Getter method for the COM property "SmartTags"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SmartTags
   */

  @VTID(206)
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

  @VTID(207)
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

  @VTID(207)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void speak(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakDirection);

  /**
   * @param speakDirection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param speakFormulas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(207)
  void speak(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakDirection,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakFormulas);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(208)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteType.class, com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4104", "-4142", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(208)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPasteSpecialOperation.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4142", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(208)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(208)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object pasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object skipBlanks);

  /**
   * @param paste Optional parameter. Default value is -4104
   * @param operation Optional parameter. Default value is -4142
   * @param skipBlanks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param transpose Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(208)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object pasteSpecial(
    @Optional @DefaultValue("-4104") com.exceljava.com4j.excel.XlPasteType paste,
    @Optional @DefaultValue("-4142") com.exceljava.com4j.excel.XlPasteSpecialOperation operation,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object skipBlanks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object transpose);


  /**
   * <p>
   * Getter method for the COM property "AllowEdit"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(209)
  boolean getAllowEdit();


  /**
   * <p>
   * Getter method for the COM property "ListObject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(210)
  com.exceljava.com4j.excel.ListObject getListObject();


  /**
   * <p>
   * Getter method for the COM property "XPath"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XPath
   */

  @VTID(211)
  com.exceljava.com4j.excel.XPath getXPath();


  /**
   * <p>
   * Getter method for the COM property "ServerActions"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Actions
   */

  @VTID(212)
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

  @VTID(213)
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

  @VTID(213)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"2"})
  void removeDuplicates(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columns);

  /**
   * @param columns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param header Optional parameter. Default value is 2
   */

  @VTID(213)
  void removeDuplicates(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columns,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {8}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
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
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(214)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object printOut(
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
   * Getter method for the COM property "MDX"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(215)
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

  @VTID(216)
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

  @VTID(216)
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

  @VTID(216)
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

  @VTID(216)
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

  @VTID(216)
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

  @VTID(216)
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

  @VTID(216)
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

  @VTID(216)
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

  @VTID(216)
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
   * Getter method for the COM property "CountLarge"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(217)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCountLarge();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(218)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object calculateRowMajorOrder();


  /**
   * <p>
   * Getter method for the COM property "SparklineGroups"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparklineGroups
   */

  @VTID(219)
  com.exceljava.com4j.excel.SparklineGroups getSparklineGroups();


  /**
   */

  @VTID(220)
  void clearHyperlinks();


  /**
   * <p>
   * Getter method for the COM property "DisplayFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DisplayFormat
   */

  @VTID(221)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(222)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "2", "-4105", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(222)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlBorderWeight.class, com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "-4105", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object borderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(222)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlColorIndex.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4105", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object borderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(222)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object borderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(222)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object borderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object color);

  /**
   * @param lineStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param weight Optional parameter. Default value is 2
   * @param colorIndex Optional parameter. Default value is -4105
   * @param color Optional parameter. Default value is com4j.Variant.getMissing()
   * @param themeColor Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(222)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object borderAround(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lineStyle,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlBorderWeight weight,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlColorIndex colorIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object color,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object themeColor);


  /**
   */

  @VTID(223)
  void allocateChanges();


  /**
   */

  @VTID(224)
  void discardChanges();


  /**
   */

  @VTID(225)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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

  @VTID(226)
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
   * Getter method for the COM property "HasRichDataType"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(227)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHasRichDataType();


  /**
   */

  @VTID(228)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.CommentThreaded
   */

  @VTID(229)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.CommentThreaded addCommentThreaded();

  /**
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CommentThreaded
   */

  @VTID(229)
  com.exceljava.com4j.excel.CommentThreaded addCommentThreaded(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object text);


  /**
   * <p>
   * Getter method for the COM property "CommentThreaded"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CommentThreaded
   */

  @VTID(230)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {16}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 16}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 16}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 16}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 16}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 16}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 16}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrder.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 16}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 16}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 16}, optParamIndex = {9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 16}, optParamIndex = {10, 11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortOrientation.class, com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"2", "1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 16}, optParamIndex = {11, 12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortMethod.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 16}, optParamIndex = {12, 13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 16}, optParamIndex = {13, 14, 15}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16}, optParamIndex = {14, 15}, javaType = {com.exceljava.com4j.excel.XlSortDataOption.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16}, optParamIndex = {15}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=16)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(231)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object sort(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object key3,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortOrder order3,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlYesNoGuess header,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object orderCustom,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlSortOrientation orientation,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSortMethod sortMethod,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption1,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption2,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlSortDataOption dataOption3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subField1);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(232)
  @UseDefaultValues(paramIndexMapping = {6}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(232)
  @UseDefaultValues(paramIndexMapping = {0, 6}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "1", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object autoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(232)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlAutoFilterOperator.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object autoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(232)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object autoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1,
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(232)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object autoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria2);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(232)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=6)
  java.lang.Object autoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visibleDropDown);

  /**
   * @param field Optional parameter. Default value is com4j.Variant.getMissing()
   * @param criteria1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is 1
   * @param criteria2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visibleDropDown Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subField Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(232)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object autoFilter(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria1,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAutoFilterOperator operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object criteria2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visibleDropDown,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subField);


  /**
   * @param serviceID Mandatory int parameter.
   * @param languageCulture Mandatory java.lang.String parameter.
   */

  @VTID(233)
  void convertToLinkedDataType(
    int serviceID,
    java.lang.String languageCulture);


  /**
   * <p>
   * Getter method for the COM property "LinkedDataTypeState"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(234)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getLinkedDataTypeState();


  /**
   * @param sourceCell Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(235)
  void setCellDataTypeFromCell(
    com.exceljava.com4j.excel.Range sourceCell);


  /**
   */

  @VTID(236)
  void dataTypeToText();


  /**
   * <p>
   * Getter method for the COM property "HasSpill"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(237)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHasSpill();


  /**
   * <p>
   * Getter method for the COM property "SpillingToRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(238)
  com.exceljava.com4j.excel.Range getSpillingToRange();


  /**
   * <p>
   * Getter method for the COM property "SpillParent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(239)
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

  @VTID(240)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void refreshLinkedDataType();

  /**
   * @param domainID Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(240)
  void refreshLinkedDataType(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object domainID);


  /**
   * <p>
   * Getter method for the COM property "Formula2"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFormula2(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(241)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getFormula2();

  /**
   * <p>
   * Getter method for the COM property "Formula2"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(241)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormula2(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Formula2"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setFormula2(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(242)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setFormula2(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Formula2"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(242)
  void setFormula2(
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula2Local"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(243)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormula2Local();


  /**
   * <p>
   * Setter method for the COM property "Formula2Local"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(244)
  void setFormula2Local(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula2R1C1"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFormula2R1C1(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(245)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getFormula2R1C1();

  /**
   * <p>
   * Getter method for the COM property "Formula2R1C1"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(245)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormula2R1C1(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Formula2R1C1"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setFormula2R1C1(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(246)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setFormula2R1C1(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Formula2R1C1"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(246)
  void setFormula2R1C1(
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula2R1C1Local"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(247)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormula2R1C1Local();


  /**
   * <p>
   * Setter method for the COM property "Formula2R1C1Local"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(248)
  void setFormula2R1C1Local(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SavedAsArray"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(249)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  @UseDefaultValues(paramIndexMapping = {0, 1, 9}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement);

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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 9}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt);

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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 9}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder);

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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 9}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase);

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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 9}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte);

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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 9}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=9)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchFormat);

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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 9}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=9)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replaceFormat);

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
   * @return  Returns a value of type boolean
   */

  @VTID(250)
  boolean replace(
    @MarshalAs(NativeType.VARIANT) java.lang.Object what,
    @MarshalAs(NativeType.VARIANT) java.lang.Object replacement,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lookAt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchCase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object matchByte,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object searchFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replaceFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formulaVersion);


  // Properties:
}
