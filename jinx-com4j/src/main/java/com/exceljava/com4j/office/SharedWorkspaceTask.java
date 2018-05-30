package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0379-0000-0000-C000-000000000046}")
public interface SharedWorkspaceTask extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  java.lang.String getTitle();


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param title Mandatory java.lang.String parameter.
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(10)
  @DefaultMethod
  void setTitle(
    java.lang.String title);


  /**
   * <p>
   * Getter method for the COM property "AssignedTo"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getAssignedTo();


  /**
   * <p>
   * Setter method for the COM property "AssignedTo"
   * </p>
   * @param assignedTo Mandatory java.lang.String parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(12)
  void setAssignedTo(
    java.lang.String assignedTo);


  /**
   * <p>
   * Getter method for the COM property "Status"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoSharedWorkspaceTaskStatus
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.MsoSharedWorkspaceTaskStatus getStatus();


  /**
   * <p>
   * Setter method for the COM property "Status"
   * </p>
   * @param status Mandatory com.exceljava.com4j.office.MsoSharedWorkspaceTaskStatus parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(14)
  void setStatus(
    com.exceljava.com4j.office.MsoSharedWorkspaceTaskStatus status);


  /**
   * <p>
   * Getter method for the COM property "Priority"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoSharedWorkspaceTaskPriority
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.MsoSharedWorkspaceTaskPriority getPriority();


  /**
   * <p>
   * Setter method for the COM property "Priority"
   * </p>
   * @param priority Mandatory com.exceljava.com4j.office.MsoSharedWorkspaceTaskPriority parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(16)
  void setPriority(
    com.exceljava.com4j.office.MsoSharedWorkspaceTaskPriority priority);


  /**
   * <p>
   * Getter method for the COM property "Description"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(17)
  java.lang.String getDescription();


  /**
   * <p>
   * Setter method for the COM property "Description"
   * </p>
   * @param description Mandatory java.lang.String parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(18)
  void setDescription(
    java.lang.String description);


  /**
   * <p>
   * Getter method for the COM property "DueDate"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(19)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDueDate();


  /**
   * <p>
   * Setter method for the COM property "DueDate"
   * </p>
   * @param dueDate Mandatory java.lang.Object parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(20)
  void setDueDate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object dueDate);


  /**
   * <p>
   * Getter method for the COM property "CreatedBy"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(21)
  java.lang.String getCreatedBy();


  /**
   * <p>
   * Getter method for the COM property "CreatedDate"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCreatedDate();


  /**
   * <p>
   * Getter method for the COM property "ModifiedBy"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(23)
  java.lang.String getModifiedBy();


  /**
   * <p>
   * Getter method for the COM property "ModifiedDate"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(24)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getModifiedDate();


  /**
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(25)
  void save();


  /**
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(26)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(27)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  // Properties:
}
