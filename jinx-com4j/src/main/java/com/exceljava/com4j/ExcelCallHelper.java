package com.exceljava.com4j;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Helper for registering actions to be run at the end of the current
 * Excel call, using com.exceljava.jinx.ExcelCall.
 *
 * ExcelCall was added in Jinx 3.0, so it is accessed using reflection
 * to avoid requiring that version. When it is not available, or when
 * there is no current call, the action is not registered.
 */
class ExcelCallHelper {
    private static final Logger log = Logger.getLogger(ExcelCallHelper.class.getName());

    private static final Method currentMethod;
    private static final Method onCloseMethod;

    static {
        Method current = null;
        Method onClose = null;
        try {
            Class<?> cls = Class.forName("com.exceljava.jinx.ExcelCall");
            current = cls.getMethod("current");
            onClose = cls.getMethod("onClose", Runnable.class);
        } catch (ClassNotFoundException | NoSuchMethodException e) {
            // ExcelCall is only available from Jinx 3.0
            current = null;
            onClose = null;
        }
        currentMethod = current;
        onCloseMethod = onClose;
    }

    /**
     * Registers an action to be run once the current Excel call has completed.
     *
     * @param action Action to run once the current call has completed.
     * @return True if the action was registered, false if ExcelCall is not
     *         available or there is no current call.
     */
    public static boolean onClose(Runnable action) {
        if (null == currentMethod || null == onCloseMethod) {
            return false;
        }

        try {
            Object call = currentMethod.invoke(null);
            if (null == call) {
                return false;
            }

            onCloseMethod.invoke(call, action);
            return true;
        } catch (IllegalAccessException | InvocationTargetException e) {
            log.log(Level.FINE, "Unable to register ExcelCall close action.", e);
            return false;
        }
    }
}
