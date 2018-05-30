package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C030A-0000-0000-C000-000000000046}")
public interface CommandBarPopup extends com.exceljava.com4j.office.CommandBarControl {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "CommandBar"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBar
   */

  @DISPID(1610940416) //= 0x60050000. The runtime will prefer the VTID if present
  @VTID(83)
  com.exceljava.com4j.office.CommandBar getCommandBar();


  /**
   * <p>
   * Getter method for the COM property "Controls"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControls
   */

  @DISPID(1610940417) //= 0x60050001. The runtime will prefer the VTID if present
  @VTID(84)
  com.exceljava.com4j.office.CommandBarControls getControls();


  @VTID(84)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CommandBarControls.class})
  com.exceljava.com4j.office.CommandBarControl getControls(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "OLEMenuGroup"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoOLEMenuGroup
   */

  @DISPID(1610940418) //= 0x60050002. The runtime will prefer the VTID if present
  @VTID(85)
  com.exceljava.com4j.office.MsoOLEMenuGroup getOLEMenuGroup();


  /**
   * <p>
   * Setter method for the COM property "OLEMenuGroup"
   * </p>
   * @param pomg Mandatory com.exceljava.com4j.office.MsoOLEMenuGroup parameter.
   */

  @DISPID(1610940418) //= 0x60050002. The runtime will prefer the VTID if present
  @VTID(86)
  void setOLEMenuGroup(
    com.exceljava.com4j.office.MsoOLEMenuGroup pomg);


  /**
   * <p>
   * Getter method for the COM property "InstanceIdPtr"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610940420) //= 0x60050004. The runtime will prefer the VTID if present
  @VTID(87)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getInstanceIdPtr();


  // Properties:
}
