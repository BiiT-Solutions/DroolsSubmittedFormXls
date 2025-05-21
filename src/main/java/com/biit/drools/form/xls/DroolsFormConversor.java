package com.biit.drools.form.xls;

import com.biit.drools.form.DroolsSubmittedCategory;
import com.biit.drools.form.DroolsSubmittedForm;
import com.biit.drools.form.DroolsSubmittedQuestion;
import com.biit.form.result.FormResult;
import com.biit.drools.form.xls.logger.XlsExporterLog;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DroolsFormConversor {
    private static final String[] COLUMNS_NAMES = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
            "V", "W", "X", "Y", "Z"};
    private static final String QUESTION_LABEL_TITLE = "Question";
    private static final int TITLE_FONT_SIZE = 14;
    private static final int ANSWER_LABEL_FONT_SIZE = 12;
    private static final double DEFAULT_ROW_EIGHT = (20 * ANSWER_LABEL_FONT_SIZE * 1.5);
    private static final int QUESTION_LABEL_WIDTH = 256 * 50;

    private static final byte GREY_50_PERCENT = (byte) 0xA0;
    private static final byte GREY_25_PERCENT = (byte) 0xEE;
    private static final byte RED_R = (byte) 0xEE;
    private static final byte RED_G = (byte) 0xAA;
    private static final byte RED_B = (byte) 0xAA;
    private static final byte PINK_R = (byte) 0xF2;
    private static final byte PINK_G = (byte) 0x0D;
    private static final byte PINK_B = (byte) 0x5E;

    private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final String NO_DATA = "         ";

    private static final int FORM_NUMBER_3 = 3;

    private static final int QUESTION_LABEL_COLUMN = 1;
    private static final int TITLE_ROW = 1;
    private static final String ANSWER_SEPARATOR = ", ";

    private final Map<String, Map<String, Integer>> questionRowNumber = new HashMap<>();
    private final Map<String, HSSFRow> questionRow = new HashMap<>();
    private final Map<String, HSSFSheet> categorySheet = new HashMap<>();

    private HSSFCellStyle titleStyle = null;
    private HSSFCellStyle answerStyle = null;
    private HSSFCellStyle contentStyle = null;

    public void createXlsDocument(HSSFWorkbook workbook, List<DroolsSubmittedForm> droolsSubmittedForms, List<String> formHeaders) {
        // Override colors
        setColor(workbook, HSSFColorPredefined.GREY_50_PERCENT, GREY_50_PERCENT, GREY_50_PERCENT, GREY_50_PERCENT);
        setColor(workbook, HSSFColorPredefined.GREY_25_PERCENT, GREY_25_PERCENT, GREY_25_PERCENT, GREY_25_PERCENT);
        setColor(workbook, HSSFColorPredefined.RED, RED_R, RED_G, RED_B);
        setColor(workbook, HSSFColorPredefined.PINK, PINK_R, PINK_G, PINK_B);

        createAnswersTables(workbook, droolsSubmittedForms, formHeaders);
    }

    public static String getColumnName(int index) {
        if (index < COLUMNS_NAMES.length) {
            return COLUMNS_NAMES[index];
        } else {
            return COLUMNS_NAMES[index / COLUMNS_NAMES.length - 1] + COLUMNS_NAMES[index % COLUMNS_NAMES.length];
        }
    }

    private void setColorToCell(HSSFSheet sheet, int totalForms, int totalQuestions) {
        if (totalForms > 1) {
            // First values
            final ConditionalFormattingRule ruleFirstColumn = sheet.getSheetConditionalFormatting().createConditionalFormattingRule(
                    "((indirect(address(row(), column() + 1))) <> (indirect(address(row(), column()))))");
            PatternFormatting fill = ruleFirstColumn.createPatternFormatting();
            fill.setFillBackgroundColor(HSSFColorPredefined.RED.getIndex());
            fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

            ConditionalFormattingRule[] cfRules = new ConditionalFormattingRule[]{ruleFirstColumn};

            CellRangeAddress[] regions = new CellRangeAddress[]{CellRangeAddress.valueOf(getColumnName(getFormScoreColumn(totalForms, 1)) + (TITLE_ROW + 2)
                    + ":" + getColumnName(getFormScoreColumn(totalForms, 1)) + (TITLE_ROW + totalQuestions + 1))};
            sheet.getSheetConditionalFormatting().addConditionalFormatting(regions, cfRules);

            // Last values
            final ConditionalFormattingRule ruleLastColum = sheet.getSheetConditionalFormatting().createConditionalFormattingRule(
                    "((indirect(address(row(), column() - 1))) <> (indirect(address(row(), column()))))");
            fill = ruleLastColum.createPatternFormatting();
            fill.setFillBackgroundColor(HSSFColorPredefined.RED.getIndex());
            fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

            cfRules = new ConditionalFormattingRule[]{ruleLastColum};

            regions = new CellRangeAddress[]{CellRangeAddress.valueOf(getColumnName(getFormScoreColumn(totalForms, totalForms)) + (TITLE_ROW + 2) + ":"
                    + getColumnName(getFormScoreColumn(totalForms, totalForms)) + (TITLE_ROW + totalQuestions + 1))};
            sheet.getSheetConditionalFormatting().addConditionalFormatting(regions, cfRules);
        }

        // Intermediate values
        if (totalForms > 2) {
            final ConditionalFormattingRule ruleIntermediateColumn = sheet
                    .getSheetConditionalFormatting()
                    .createConditionalFormattingRule(
                            "OR(((indirect(address(row(), column() - 1))) <> (indirect(address(row(), column())))), "
                                    + "((indirect(address(row(), column() + 1))) <> (indirect(address(row(), column()))))) ");

            final PatternFormatting fill = ruleIntermediateColumn.createPatternFormatting();
            fill.setFillBackgroundColor(HSSFColorPredefined.RED.getIndex());
            fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

            final ConditionalFormattingRule[] cfRules = new ConditionalFormattingRule[]{ruleIntermediateColumn};

            final CellRangeAddress[] regions = new CellRangeAddress[]{CellRangeAddress.valueOf(getColumnName(getFormScoreColumn(totalForms, 2))
                    + (TITLE_ROW + 2) + ":" + getColumnName(getFormScoreColumn(totalForms, totalForms - 1)) + (TITLE_ROW + totalQuestions + 1))};
            sheet.getSheetConditionalFormatting().addConditionalFormatting(regions, cfRules);
        }
    }

    private void createAnswersTables(HSSFWorkbook workbook, List<DroolsSubmittedForm> droolsSubmittedForms, List<String> formHeaders) {
        for (int i = 0; i < droolsSubmittedForms.size(); i++) {
            for (DroolsSubmittedCategory category : droolsSubmittedForms.get(i).getAllChildrenInHierarchy(DroolsSubmittedCategory.class)) {
                if (!category.getAllChildrenInHierarchy(DroolsSubmittedQuestion.class).isEmpty()) {
                    final HSSFSheet sheet = getSheet(workbook, droolsSubmittedForms.get(i), category);

                    // Create title
                    createAnswersTitle(workbook, sheet, droolsSubmittedForms, formHeaders);

                    // Create answer rows
                    for (DroolsSubmittedQuestion child : category.getAllChildrenInHierarchy(DroolsSubmittedQuestion.class)) {
                        createRow(workbook, sheet, category, child, i + 1);
                    }
                    try {
                        sheet.autoSizeColumn(getFormResultColumn(i + 1));
                        sheet.autoSizeColumn(QUESTION_LABEL_COLUMN);
                    } catch (NullPointerException e) {
                        // Font not available.
                    }
                }
            }
        }
    }

    private HSSFSheet getSheet(HSSFWorkbook workbook, DroolsSubmittedForm droolsSubmittedForm, DroolsSubmittedCategory category) {
        if (categorySheet.get(droolsSubmittedForm.getText() + "_" + category.getName()) == null) {
            final HSSFSheet sheet = workbook.createSheet(parseInvalidCharacters(category.getText()));
            sheet.setDefaultRowHeight((short) DEFAULT_ROW_EIGHT);
            categorySheet.put(droolsSubmittedForm.getText() + "_" + category.getName(), sheet);
        }
        return categorySheet.get(droolsSubmittedForm.getText() + "_" + category.getName());
    }

    private String parseInvalidCharacters(String text) {
        // Sheets does not allows this characters.
        return text.replace(":", "").replace("\\", "-").replace("/", "-").replace("*", "").replace("?", "").replace("[", "(").replace("]", ")");
    }

    private void createAnswersTitle(HSSFWorkbook workbook, HSSFSheet sheet, List<DroolsSubmittedForm> forms, List<String> formHeaders) {
        final HSSFRow titleRow = sheet.createRow(TITLE_ROW);
        titleRow.createCell(QUESTION_LABEL_COLUMN).setCellValue(QUESTION_LABEL_TITLE);
        titleRow.getCell(QUESTION_LABEL_COLUMN).setCellStyle(getTitleStyle(workbook));
        // sheet.autoSizeColumn(QUESTION_LABEL_COLUMN);
        sheet.setColumnWidth(QUESTION_LABEL_COLUMN, QUESTION_LABEL_WIDTH);

        for (int i = 0; i < forms.size(); i++) {
            if (formHeaders != null && i < formHeaders.size()) {
                titleRow.createCell(getFormResultColumn(i) + 1).setCellValue(formHeaders.get(i));
            } else if (forms.get(i).getSubmittedBy() != null) {
                titleRow.createCell(getFormResultColumn(i) + 1).setCellValue(forms.get(i).getSubmittedBy());
            } else if (forms.get(i).getSubmittedAt() != null) {
                titleRow.createCell(getFormResultColumn(i) + 1).setCellValue(forms.get(i).getSubmittedAt().format(DATE_TIME_FORMATTER));
            } else {
                titleRow.createCell(getFormResultColumn(i) + 1).setCellValue(NO_DATA);
            }
            titleRow.getCell(getFormResultColumn(i) + 1).setCellStyle(getTitleStyle(workbook));
            try {
                sheet.autoSizeColumn(getFormResultColumn(i) + 1);
            } catch (NullPointerException e) {
                // Font not available.
            }
        }
    }

    private void createScoreTitle(HSSFWorkbook workbook, HSSFSheet sheet, List<FormResult> forms, List<String> formHeaders) {
        final HSSFRow titleRow = sheet.getRow(TITLE_ROW);

        for (int i = 0; i < forms.size(); i++) {
            if (i < formHeaders.size()) {
                titleRow.createCell(getFormScoreColumn(forms.size(), i) + 1).setCellValue(formHeaders.get(i));
            } else {
                titleRow.createCell(getFormScoreColumn(forms.size(), i) + 1).setCellValue(forms.get(i).getName());
            }
            titleRow.getCell(getFormScoreColumn(forms.size(), i) + 1).setCellStyle(getTitleStyle(workbook));
            try {
                sheet.autoSizeColumn(getFormScoreColumn(forms.size(), i) + 1);
            } catch (NullPointerException e) {
                // Font not available.
            }
            // sheet.setColumnWidth(QUESTION_LABEL_COLUMN + formNumber, 2000);
        }
    }

    private void createSummatoryScoreTitle(HSSFWorkbook workbook, HSSFSheet sheet, List<FormResult> forms, List<String> formHeaders) {
        final HSSFRow titleRow = sheet.getRow(TITLE_ROW);
        titleRow.createCell(getSummatoryScoreColumn(forms.size())).setCellValue("Summatory");
        titleRow.getCell(getSummatoryScoreColumn(forms.size())).setCellStyle(getTitleStyle(workbook));

        try {
            sheet.autoSizeColumn(getSummatoryScoreColumn(forms.size()));
        } catch (NullPointerException e) {
            // Font not available.
        }
    }

    private void createRow(HSSFWorkbook workbook, HSSFSheet sheet, DroolsSubmittedCategory category, DroolsSubmittedQuestion question, int formNumber) {
        setCellValue(workbook, sheet, category, question, getFormResultColumn(formNumber), getAnswersText(question));
    }

    private void setCellValue(HSSFWorkbook workbook, HSSFSheet sheet, DroolsSubmittedCategory category, DroolsSubmittedQuestion question, int column, String value) {
        final HSSFRow questionRow = getQuestionRow(workbook, sheet, category, question);
        questionRow.createCell(column).setCellValue(value);
        questionRow.getCell(column).setCellStyle(getContentStyle(workbook));
    }

    private void setCellFormula(HSSFWorkbook workbook, HSSFSheet sheet, DroolsSubmittedCategory category, DroolsSubmittedQuestion question, int column, String formula) {
        final HSSFRow questionRow = getQuestionRow(workbook, sheet, category, question);
        final HSSFCell cell = questionRow.createCell(column);
        cell.setCellFormula(formula);
        cell.setCellStyle(getContentStyle(workbook));
    }

    private String getAnswersText(DroolsSubmittedQuestion question) {
        // Add answers
        final List<String> answers = new ArrayList<>(question.getAnswers());
        Collections.sort(answers);
        final StringBuilder stringBuilder = new StringBuilder();
        for (String answer : answers) {
            if (answer != null) {
                stringBuilder.append(answer);
                stringBuilder.append(ANSWER_SEPARATOR);
            }
        }
        // Remove last separator
        if (stringBuilder.length() > 0) {
            return stringBuilder.substring(0, stringBuilder.length() - ANSWER_SEPARATOR.length());
        }
        return stringBuilder.toString();
    }

    private HSSFRow getQuestionRow(HSSFWorkbook workbook, HSSFSheet sheet, DroolsSubmittedCategory category, DroolsSubmittedQuestion question) {
        if (questionRow.get(question.getXPath()) == null) {
            questionRow.put(question.getXPath(), sheet.createRow(getQuestionRowNumber(category, question)));
            questionRow.get(question.getXPath()).createCell(QUESTION_LABEL_COLUMN).setCellValue(question.getText());
            questionRow.get(question.getXPath()).getCell(QUESTION_LABEL_COLUMN).setCellStyle(getAnswerLabelsStyle(workbook));
        }
        return questionRow.get(question.getXPath());
    }

    private int getQuestionRowNumber(DroolsSubmittedCategory category, DroolsSubmittedQuestion question) {
        if (questionRowNumber.get(category.getXPath()) == null) {
            questionRowNumber.put(category.getXPath(), new HashMap<>());
        }
        if (questionRowNumber.get(category.getXPath()).get(question.getXPath()) == null) {
            questionRowNumber.get(category.getXPath()).put(question.getXPath(), questionRowNumber.get(category.getXPath()).size() + TITLE_ROW + 1);
        }
        return questionRowNumber.get(category.getXPath()).get(question.getXPath());
    }

    private int getFormResultColumn(int formNumber) {
        return formNumber + QUESTION_LABEL_COLUMN;
    }

    private int getFormScoreColumn(int totalforms, int formNumber) {
        return totalforms + getFormResultColumn(formNumber) + 1;
    }

    private int getSummatoryScoreColumn(int totalforms) {
        return getFormScoreColumn(totalforms, totalforms) + 1;
    }

    private int getWeightScoreColumn(int totalforms) {
        return getSummatoryScoreColumn(totalforms) + 1;
    }

    private int getFormScoreValuesColumn(int totalforms, int formNumber) {
        return getWeightScoreColumn(totalforms) + formNumber + 1;
    }

    private int getFormWeightValuesColumn(int totalforms, int formNumber) {
        return getFormScoreValuesColumn(totalforms, formNumber) + FORM_NUMBER_3;
    }

    private CellStyle getAnswerLabelsStyle(HSSFWorkbook workbook) {
        if (answerStyle == null) {
            answerStyle = workbook.createCellStyle();

            // Background Color
            answerStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
            answerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Font
            final Font answerFont = workbook.createFont();
            answerFont.setFontHeightInPoints((short) ANSWER_LABEL_FONT_SIZE);
            answerStyle.setFont(answerFont);

            // Border
            answerStyle.setBorderBottom(BorderStyle.THIN);
            answerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        }
        return answerStyle;
    }

    private CellStyle getTitleStyle(HSSFWorkbook workbook) {
        if (titleStyle == null) {
            titleStyle = workbook.createCellStyle();

            // Border
            titleStyle.setBorderRight(BorderStyle.THIN);
            titleStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            titleStyle.setBorderBottom(BorderStyle.THIN);
            titleStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            titleStyle.setBorderLeft(BorderStyle.THIN);
            titleStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            titleStyle.setBorderTop(BorderStyle.THIN);
            titleStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

            // Background Color
            titleStyle.setFillForegroundColor(HSSFColorPredefined.PINK.getIndex());
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Alignment
            titleStyle.setAlignment(HorizontalAlignment.CENTER);
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // Font
            final Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) TITLE_FONT_SIZE);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            titleStyle.setFont(headerFont);
        }
        return titleStyle;
    }

    private CellStyle getContentStyle(HSSFWorkbook workbook) {
        if (contentStyle == null) {
            contentStyle = workbook.createCellStyle();

            // Background Color
            contentStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
            contentStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Border
            contentStyle.setBorderBottom(BorderStyle.THIN);
            contentStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        }
        return contentStyle;
    }

    private Collection<DroolsSubmittedCategory> getUniqueCategories(List<DroolsSubmittedForm> formResults) {
        final Map<String, DroolsSubmittedCategory> categories = new HashMap<>();
        for (DroolsSubmittedForm formResult : formResults) {
            for (DroolsSubmittedCategory category : formResult.getAllChildrenInHierarchy(DroolsSubmittedCategory.class)) {
                categories.putIfAbsent(category.getName(), category);
            }
        }
        return categories.values();
    }

    private HSSFColor setColor(HSSFWorkbook workbook, HSSFColorPredefined color, byte r, byte g, byte b) {
        final HSSFPalette palette = workbook.getCustomPalette();
        HSSFColor hssfColor = null;
        try {
            hssfColor = palette.findColor(r, g, b);
            if (hssfColor == null) {
                palette.setColorAtIndex(color.getIndex(), r, g, b);
                hssfColor = palette.getColor(color.getIndex());
            }
        } catch (Exception e) {
            XlsExporterLog.errorMessage(this.getClass().getName(), e);
        }

        return hssfColor;
    }
}
