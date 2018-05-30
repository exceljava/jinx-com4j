package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03E4-0000-0000-C000-000000000046}")
public interface PickerResult extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Id"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String getId();


  /**
   * <p>
   * Getter method for the COM property "DisplayName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getDisplayName();


  /**
   * <p>
   * Setter method for the COM property "DisplayName"
   * </p>
   * @param displayName Mandatory java.lang.String parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  void setDisplayName(
    java.lang.String displayName);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param type Mandatory java.lang.String parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  void setType(
    java.lang.String type);


  /**
   * <p>
   * Getter method for the COM property "SIPId"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  java.lang.String getSIPId();


  /**
   * <p>
   * Setter method for the COM property "SIPId"
   * </p>
   * @param sipId Mandatory java.lang.String parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(15)
  void setSIPId(
    java.lang.String sipId);


  /**
   * <p>
   * Getter method for the COM property "ItemData"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(16)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getItemData();


  /**
   * <p>
   * Setter method for the COM property "ItemData"
   * </p>
   * @param itemData Mandatory java.lang.Object parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(17)
  void setItemData(
    @MarshalAs(NativeType.VARIANT) java.lang.Object itemData);


  /**
   * <p>
   * Getter method for the COM property "SubItems"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(18)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSubItems();


  /**
   * <p>
   * Setter method for the COM property "SubItems"
   * </p>
   * @param subItems Mandatory java.lang.Object parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(19)
  void setSubItems(
    @MarshalAs(NativeType.VARIANT) java.lang.Object subItems);


  /**
   * <p>
   * Getter method for the COM property "DuplicateResults"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(20)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDuplicateResults();


  /**
   * <p>
   * Getter method for the COM property "Fields"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.PickerFields
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.PickerFields getFields();


  @VTID(21)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.PickerFields.class})
  com.exceljava.com4j.office.PickerField getFields(
    int index);

  /**
   * <p>
   * Setter method for the COM property "Fields"
   * </p>
   * @param fields Mandatory com.exceljava.com4j.office.PickerFields parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(22)
  void setFields(
    com.exceljava.com4j.office.PickerFields fields);


  // Properties:
}
