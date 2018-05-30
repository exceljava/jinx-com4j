package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0385-0000-0000-C000-000000000046}")
public interface SharedWorkspace extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param name Mandatory java.lang.String parameter.
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(10)
  @DefaultMethod
  void setName(
    java.lang.String name);


  /**
   * <p>
   * Getter method for the COM property "Members"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceMembers
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.SharedWorkspaceMembers getMembers();


  @VTID(11)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SharedWorkspaceMembers.class})
  com.exceljava.com4j.office.SharedWorkspaceMember getMembers(
    int index);

  /**
   * <p>
   * Getter method for the COM property "Tasks"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceTasks
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.SharedWorkspaceTasks getTasks();


  @VTID(12)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SharedWorkspaceTasks.class})
  com.exceljava.com4j.office.SharedWorkspaceTask getTasks(
    int index);

  /**
   * <p>
   * Getter method for the COM property "Files"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceFiles
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.SharedWorkspaceFiles getFiles();


  @VTID(13)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SharedWorkspaceFiles.class})
  com.exceljava.com4j.office.SharedWorkspaceFile getFiles(
    int index);

  /**
   * <p>
   * Getter method for the COM property "Folders"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceFolders
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.SharedWorkspaceFolders getFolders();


  @VTID(14)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SharedWorkspaceFolders.class})
  com.exceljava.com4j.office.SharedWorkspaceFolder getFolders(
    int index);

  /**
   * <p>
   * Getter method for the COM property "Links"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceLinks
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.SharedWorkspaceLinks getLinks();


  @VTID(15)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SharedWorkspaceLinks.class})
  com.exceljava.com4j.office.SharedWorkspaceLink getLinks(
    int index);

  /**
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(16)
  void refresh();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter url is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createNew(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void createNew();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createNew(url, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param url Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void createNew(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object url);

  /**
   * @param url Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(17)
  void createNew(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object url,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);


  /**
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(18)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(19)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "URL"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(20)
  java.lang.String getURL();


  /**
   * <p>
   * Getter method for the COM property "Connected"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(21)
  boolean getConnected();


  /**
   * <p>
   * Getter method for the COM property "LastRefreshed"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getLastRefreshed();


  /**
   * <p>
   * Getter method for the COM property "SourceURL"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(23)
  java.lang.String getSourceURL();


  /**
   * <p>
   * Setter method for the COM property "SourceURL"
   * </p>
   * @param pbstrSourceURL Mandatory java.lang.String parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(24)
  void setSourceURL(
    java.lang.String pbstrSourceURL);


  /**
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(25)
  void removeDocument();


  /**
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(26)
  void disconnect();


  // Properties:
}
