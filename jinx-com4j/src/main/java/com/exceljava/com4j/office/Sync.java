package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0386-0000-0000-C000-000000000046}")
public interface Sync extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Status"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoSyncStatusType
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.MsoSyncStatusType getStatus();


  /**
   * <p>
   * Getter method for the COM property "WorkspaceLastChangedBy"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getWorkspaceLastChangedBy();


  /**
   * <p>
   * Getter method for the COM property "LastSyncTime"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getLastSyncTime();


  /**
   * <p>
   * Getter method for the COM property "ErrorType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoSyncErrorType
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.MsoSyncErrorType getErrorType();


  /**
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(13)
  void getUpdate();


  /**
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(14)
  void putUpdate();


  /**
   * @param syncVersionType Mandatory com.exceljava.com4j.office.MsoSyncVersionType parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(15)
  void openVersion(
    com.exceljava.com4j.office.MsoSyncVersionType syncVersionType);


  /**
   * @param syncConflictResolution Mandatory com.exceljava.com4j.office.MsoSyncConflictResolutionType parameter.
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(16)
  void resolveConflict(
    com.exceljava.com4j.office.MsoSyncConflictResolutionType syncConflictResolution);


  /**
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(17)
  void unsuspend();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(18)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  // Properties:
}
