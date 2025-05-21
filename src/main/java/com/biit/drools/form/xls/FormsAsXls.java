package com.biit.drools.form.xls;

import com.biit.drools.form.DroolsSubmittedForm;
import com.biit.drools.form.xls.exceptions.InvalidXlsElementException;
import com.biit.drools.form.xls.logger.XlsExporterLog;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class FormsAsXls {
    private List<DroolsSubmittedForm> droolsSubmittedForms;
    private List<String> formHeaders;

    public FormsAsXls(DroolsSubmittedForm droolsSubmittedForm, String formHeader) {
        this.droolsSubmittedForms = new ArrayList<>();
        droolsSubmittedForms.add(droolsSubmittedForm);
        this.formHeaders = new ArrayList<>();
        formHeaders.add(formHeader);
    }

    public FormsAsXls(List<DroolsSubmittedForm> droolsSubmittedForms, List<String> formHeaders) {
        this.droolsSubmittedForms = droolsSubmittedForms;
        this.formHeaders = formHeaders;
    }

    public byte[] generate() throws InvalidXlsElementException {
        try {
            final HSSFWorkbook workbook = new HSSFWorkbook();

            new DroolsFormConversor().createXlsDocument(workbook, droolsSubmittedForms, formHeaders);

            final ByteArrayOutputStream fileOut = new ByteArrayOutputStream();
            workbook.write(fileOut);
            workbook.close();

            try {
                return fileOut.toByteArray();
            } finally {
                fileOut.close();
            }
        } catch (Exception e) {
            XlsExporterLog.errorMessage(this.getClass().getName(), e);
            throw new InvalidXlsElementException(e);
        }
    }

    public void createFile(String path) throws IOException, InvalidXlsElementException {
        if (!path.endsWith(".xls")) {
            path += ".xls";
        }

        try (FileOutputStream fos = new FileOutputStream(path)) {
            fos.write(generate());
        }
    }

}
