package com.biit.drools.form.xls;

/*-
 * #%L
 * Drools Submitted Form XLS Conversor
 * %%
 * Copyright (C) 2025 BiiT Sourcing Solutions S.L.
 * %%
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 * #L%
 */

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
