package com.biit.drools.form.xls.exceptions;

public class InvalidXlsElementException extends Exception {
    private static final long serialVersionUID = -4027664690256794077L;

    public InvalidXlsElementException(String text) {
        super(text);
    }

    public InvalidXlsElementException(Throwable e) {
        super(e);
    }
}
