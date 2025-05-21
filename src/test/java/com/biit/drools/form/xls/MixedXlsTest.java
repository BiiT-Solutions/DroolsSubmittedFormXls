package com.biit.drools.form.xls;

import com.biit.drools.form.DroolsSubmittedForm;
import com.biit.drools.form.xls.exceptions.InvalidXlsElementException;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

@Test(groups = {"convertXls"})
public class MixedXlsTest {
    private static final String OUTPUT_FOLDER = System.getProperty("java.io.tmpdir") + File.separator + "XmlForms";
    private static final String FORM_AS_JSON = "";
    private static final String FORM_AS_JSON_2 = "The 5 Frustrations on Teamworking 2.json";
    private static final String FORM_AS_JSON_3 = "The 5 Frustrations on Teamworking 3.json";

    @BeforeClass
    public void prepareFolder() throws IOException {
        Files.createDirectories(Paths.get(OUTPUT_FOLDER));
    }

    @Test
    public void multipleXlsFile() throws IOException, URISyntaxException, InvalidXlsElementException {
        List<DroolsSubmittedForm> DroolsSubmittedForms = new ArrayList<>();
        List<String> formHeaders = new ArrayList<>();
        // Load form from json file in resources.
        String text = new String(Files.readAllBytes(Paths.get(getClass().getClassLoader().getResource(FORM_AS_JSON).toURI())));
        DroolsSubmittedForm form = DroolsSubmittedForm.getFromJson(text);
        Assert.assertNotNull(form);
        DroolsSubmittedForms.add(form);

        text = new String(Files.readAllBytes(Paths.get(getClass().getClassLoader().getResource(FORM_AS_JSON_2).toURI())));
        form = DroolsSubmittedForm.getFromJson(text);
        Assert.assertNotNull(form);
        DroolsSubmittedForms.add(form);

        text = new String(Files.readAllBytes(Paths.get(getClass().getClassLoader().getResource(FORM_AS_JSON_3).toURI())));
        form = DroolsSubmittedForm.getFromJson(text);
        Assert.assertNotNull(form);
        DroolsSubmittedForms.add(form);

        // Convert to xls.
        FormsAsXls xlsDocument = new FormsAsXls(DroolsSubmittedForms, formHeaders);
        xlsDocument.createFile(OUTPUT_FOLDER + File.separator + "mixed.xls");
    }

    private boolean deleteDirectory(File directoryToBeDeleted) {
        File[] allContents = directoryToBeDeleted.listFiles();
        if (allContents != null) {
            for (File file : allContents) {
                deleteDirectory(file);
            }
        }
        return directoryToBeDeleted.delete();
    }

    @AfterClass
    public void removeFolder() {
        Assert.assertTrue(deleteDirectory(new File(OUTPUT_FOLDER)));
    }
}
