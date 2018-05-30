package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C037A-0000-0000-C000-000000000046}")
public interface SharedWorkspaceTasks extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTask
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.SharedWorkspaceTask getItem(
    int index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter status is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter priority is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter assignee is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dueDate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(title, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param title Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTask
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 6}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.SharedWorkspaceTask add(
    java.lang.String title);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter priority is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter assignee is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dueDate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(title, status, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param title Mandatory java.lang.String parameter.
   * @param status Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTask
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.SharedWorkspaceTask add(
    java.lang.String title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object status);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter assignee is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dueDate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(title, status, priority, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param title Mandatory java.lang.String parameter.
   * @param status Optional parameter. Default value is com4j.Variant.getMissing()
   * @param priority Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTask
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.SharedWorkspaceTask add(
    java.lang.String title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object status,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object priority);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dueDate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(title, status, priority, assignee, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param title Mandatory java.lang.String parameter.
   * @param status Optional parameter. Default value is com4j.Variant.getMissing()
   * @param priority Optional parameter. Default value is com4j.Variant.getMissing()
   * @param assignee Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTask
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.SharedWorkspaceTask add(
    java.lang.String title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object status,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object priority,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object assignee);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dueDate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(title, status, priority, assignee, description, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param title Mandatory java.lang.String parameter.
   * @param status Optional parameter. Default value is com4j.Variant.getMissing()
   * @param priority Optional parameter. Default value is com4j.Variant.getMissing()
   * @param assignee Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTask
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.SharedWorkspaceTask add(
    java.lang.String title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object status,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object priority,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object assignee,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description);

  /**
   * @param title Mandatory java.lang.String parameter.
   * @param status Optional parameter. Default value is com4j.Variant.getMissing()
   * @param priority Optional parameter. Default value is com4j.Variant.getMissing()
   * @param assignee Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dueDate Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTask
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.SharedWorkspaceTask add(
    java.lang.String title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object status,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object priority,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object assignee,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dueDate);


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "ItemCountExceeded"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(13)
  boolean getItemCountExceeded();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
