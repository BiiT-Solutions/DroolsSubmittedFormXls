package com.biit.drools.form.xls.logger;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.Writer;

public final class XlsExporterLog {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsExporterLog.class);

    private XlsExporterLog() {

    }

    /**
     * Events that have business meaning (i.e. creating category, deleting form,
     * ...). To follow user actions.
     *
     * @param message
     */
    private static void info(String message) {
        LOGGER.info(message);
    }

    /**
     * Events that have business meaning (i.e. creating category, deleting form,
     * ...). To follow user actions.
     */
    public static void info(String className, String message) {
        info(className + ": " + message);
    }

    /**
     * Shows not critical errors. I.e. Email address not found, permissions not
     * allowed for this user, ...
     *
     * @param message
     */
    private static void warning(String message) {
        LOGGER.warn(message);
    }

    /**
     * Shows not critical errors. I.e. Email address not found, permissions not
     * allowed for this user, ...
     *
     * @param message
     */
    public static void warning(String className, String message) {
        warning(className + ": " + message);
    }

    /**
     * For following the trace of the execution. I.e. Knowing if the application
     * access to a method, opening database connection, etc.
     *
     * @param message
     */
    private static void debug(String message) {
        if (isDebugEnabled()) {
            LOGGER.debug(message);
        }
    }

    /**
     * For following the trace of the execution. I.e. Knowing if the application
     * access to a method, opening database connection, etc.
     */
    public static void debug(String className, String message) {
        debug(className + ": " + message);
    }

    /**
     * To log any not expected error that can cause application malfuncionality.
     * I.e. couldn't open database connection, etc..
     *
     * @param message
     */
    private static void severe(String message) {
        LOGGER.error(message);
    }

    /**
     * To log any not expected error that can cause application malfuncionality.
     *
     * @param message
     */
    public static void severe(String className, String message) {
        severe(className + ": " + message);
    }

    /**
     * To log java exceptions and log also the stack trace.
     *
     * @param className
     * @param throwable
     */
    public static void errorMessage(String className, Throwable throwable) {
        final String error = getStackTrace(throwable);
        severe(className, error);
    }

    private static String getStackTrace(Throwable throwable) {
        final Writer writer = new StringWriter();
        final PrintWriter printWriter = new PrintWriter(writer);
        throwable.printStackTrace(printWriter);
        return writer.toString();
    }

    public static boolean isDebugEnabled() {
        return LOGGER.isDebugEnabled();
    }
}
