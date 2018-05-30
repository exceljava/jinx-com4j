package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0341-0000-0000-C000-000000000046}")
public interface Script extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809344) //= 0x60030000. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Extended"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getExtended();


  /**
   * <p>
   * Setter method for the COM property "Extended"
   * </p>
   * @param extended Mandatory java.lang.String parameter.
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(11)
  void setExtended(
    java.lang.String extended);


  /**
   * <p>
   * Getter method for the COM property "Id"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getId();


  /**
   * <p>
   * Setter method for the COM property "Id"
   * </p>
   * @param id Mandatory java.lang.String parameter.
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(13)
  void setId(
    java.lang.String id);


  /**
   * <p>
   * Getter method for the COM property "Language"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoScriptLanguage
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.MsoScriptLanguage getLanguage();


  /**
   * <p>
   * Setter method for the COM property "Language"
   * </p>
   * @param language Mandatory com.exceljava.com4j.office.MsoScriptLanguage parameter.
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(15)
  void setLanguage(
    com.exceljava.com4j.office.MsoScriptLanguage language);


  /**
   * <p>
   * Getter method for the COM property "Location"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoScriptLocation
   */

  @DISPID(1610809351) //= 0x60030007. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.MsoScriptLocation getLocation();


  /**
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Shape"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809353) //= 0x60030009. The runtime will prefer the VTID if present
  @VTID(18)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getShape();


  /**
   * <p>
   * Getter method for the COM property "ScriptText"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(19)
  @DefaultMethod
  java.lang.String getScriptText();


  /**
   * <p>
   * Setter method for the COM property "ScriptText"
   * </p>
   * @param script Mandatory java.lang.String parameter.
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(20)
  @DefaultMethod
  void setScriptText(
    java.lang.String script);


  // Properties:
}
