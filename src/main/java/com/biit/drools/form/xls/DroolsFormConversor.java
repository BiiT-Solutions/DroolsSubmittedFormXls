package com.biit.drools.form.xls;

import com.biit.drools.form.DroolsSubmittedCategory;
import com.biit.drools.form.DroolsSubmittedForm;
import com.biit.drools.form.DroolsSubmittedQuestion;
import com.biit.drools.form.xls.logger.XlsExporterLog;
import com.biit.form.submitted.implementation.SubmittedObject;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.SortedSet;
import java.util.TreeSet;

public class DroolsFormConversor {
    private static final String[] COLUMNS_NAMES = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
            "V", "W", "X", "Y", "Z"};
    private static final String QUESTION_LABEL_TITLE = "Question";
    private static final String VARIABLE_LABEL_TITLE = "Variables";
    private static final String VARIABLE_SCOPE_TITLE = "Scope";
    private static final int TITLE_FONT_SIZE = 14;
    private static final int ANSWER_LABEL_FONT_SIZE = 12;
    private static final double DEFAULT_ROW_EIGHT = (20 * ANSWER_LABEL_FONT_SIZE * 1.5);
    private static final int QUESTION_LABEL_WIDTH = 256 * 50;
    private static final int VARIABLE_LABEL_WIDTH = 128 * 50;
    private static final String VARIABLES_SHEET_NAME = "Variables";

    private static final byte GREY_50_PERCENT = (byte) 0xA0;
    private static final byte GREY_25_PERCENT = (byte) 0xEE;
    private static final byte RED_R = (byte) 0xEE;
    private static final byte RED_G = (byte) 0xAA;
    private static final byte RED_B = (byte) 0xAA;
    private static final byte PINK_R = (byte) 0xF2;
    private static final byte PINK_G = (byte) 0x0D;
    private static final byte PINK_B = (byte) 0x5E;

    private static final NumberFormat DECIMAL_FORMAT = new DecimalFormat("##.###");
    private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final String NO_DATA = "         ";

    private static final int QUESTION_LABEL_COLUMN = 1;
    private static final int VARIABLE_LABEL_COLUMN = 2;
    private static final int VARIABLE_SCOPE_COLUMN = 1;
    private static final int TITLE_ROW = 1;
    private static final String ANSWER_SEPARATOR = ", ";

    private final Map<String, Map<String, Integer>> questionRowNumber = new HashMap<>();
    private final Map<String, Integer> variableRowNumber = new HashMap<>();
    private final Map<String, HSSFRow> questionRow = new HashMap<>();
    private final Map<String, HSSFRow> variableRow = new HashMap<>();
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

        createVariableTables(workbook, droolsSubmittedForms, formHeaders);
    }

    public static String getColumnName(int index) {
        if (index < COLUMNS_NAMES.length) {
            return COLUMNS_NAMES[index];
        } else {
            return COLUMNS_NAMES[index / COLUMNS_NAMES.length - 1] + COLUMNS_NAMES[index % COLUMNS_NAMES.length];
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

    private HSSFSheet getVariablesSheet(HSSFWorkbook workbook, DroolsSubmittedForm droolsSubmittedForm) {
        if (categorySheet.get(VARIABLES_SHEET_NAME) == null) {
            final HSSFSheet sheet = workbook.createSheet(parseInvalidCharacters(VARIABLES_SHEET_NAME));
            sheet.setDefaultRowHeight((short) DEFAULT_ROW_EIGHT);
            categorySheet.put(VARIABLES_SHEET_NAME, sheet);
        }
        return categorySheet.get(VARIABLES_SHEET_NAME);
    }

    private String parseInvalidCharacters(String text) {
        // Sheets does not allows this characters.
        return text.replace(":", "").replace("\\", "-").replace("/", "-").replace("*", "").replace("?", "").replace("[", "(").replace("]", ")");
    }

    private void createAnswersTitle(HSSFWorkbook workbook, HSSFSheet sheet, List<DroolsSubmittedForm> forms, List<String> formHeaders) {
        final HSSFRow titleRow = sheet.createRow(TITLE_ROW);
        titleRow.createCell(QUESTION_LABEL_COLUMN).setCellValue(QUESTION_LABEL_TITLE);
        titleRow.getCell(QUESTION_LABEL_COLUMN).setCellStyle(getTitleStyle(workbook));
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


    private void createVariableTables(HSSFWorkbook workbook, List<DroolsSubmittedForm> droolsSubmittedForms, List<String> formHeaders) {
        for (int i = 0; i < droolsSubmittedForms.size(); i++) {
            if (droolsSubmittedForms.get(i).getVariablesValue() != null) {
                final HSSFSheet sheet = getVariablesSheet(workbook, droolsSubmittedForms.get(i));

                // Create title
                createVariablesTitle(workbook, sheet, droolsSubmittedForms, formHeaders);

                //Sort the xpaths
                final SortedSet<String> xpaths = new TreeSet<>(droolsSubmittedForms.get(i).getFormVariables().keySet());

                //Sort the variables
                for (String xpath : xpaths) {
                    final SortedSet<String> keys = new TreeSet<>(droolsSubmittedForms.get(i).getFormVariables().get(xpath).keySet());

                    final SubmittedObject scopeElement = droolsSubmittedForms.get(i).getElement(xpath);

                    // Set the variables
                    for (String key : keys) {
                        createVariableRow(workbook, sheet,
                                xpath,
                                key, String.valueOf(droolsSubmittedForms.get(i).getFormVariables().get(xpath).get(key)),
                                (scopeElement != null ? scopeElement.getText() : droolsSubmittedForms.get(i).getName()),
                                i + 1);
                    }
                }
            }
        }
    }

    private void createVariablesTitle(HSSFWorkbook workbook, HSSFSheet sheet, List<DroolsSubmittedForm> forms, List<String> formHeaders) {
        final HSSFRow titleRow = sheet.createRow(TITLE_ROW);

        titleRow.createCell(VARIABLE_SCOPE_COLUMN).setCellValue(VARIABLE_SCOPE_TITLE);
        titleRow.getCell(VARIABLE_SCOPE_COLUMN).setCellStyle(getTitleStyle(workbook));
        sheet.setColumnWidth(VARIABLE_SCOPE_COLUMN, VARIABLE_LABEL_WIDTH);

        titleRow.createCell(VARIABLE_LABEL_COLUMN).setCellValue(VARIABLE_LABEL_TITLE);
        titleRow.getCell(VARIABLE_LABEL_COLUMN).setCellStyle(getTitleStyle(workbook));
        sheet.setColumnWidth(VARIABLE_LABEL_COLUMN, VARIABLE_LABEL_WIDTH);

        for (int i = 0; i < forms.size(); i++) {
            if (formHeaders != null && i < formHeaders.size()) {
                titleRow.createCell(getFormResultColumn(i) + VARIABLE_LABEL_COLUMN).setCellValue(formHeaders.get(i));
            } else if (forms.get(i).getSubmittedBy() != null) {
                titleRow.createCell(getFormResultColumn(i) + VARIABLE_LABEL_COLUMN).setCellValue(forms.get(i).getSubmittedBy());
            } else if (forms.get(i).getSubmittedAt() != null) {
                titleRow.createCell(getFormResultColumn(i) + VARIABLE_LABEL_COLUMN).setCellValue(forms.get(i).getSubmittedAt().format(DATE_TIME_FORMATTER));
            } else {
                titleRow.createCell(getFormResultColumn(i) + VARIABLE_LABEL_COLUMN).setCellValue(NO_DATA);
            }
            titleRow.getCell(getFormResultColumn(i) + VARIABLE_LABEL_COLUMN).setCellStyle(getTitleStyle(workbook));
            try {
                sheet.autoSizeColumn(getFormResultColumn(i) + VARIABLE_LABEL_COLUMN);
            } catch (NullPointerException e) {
                // Font not available.
            }
        }
    }

    private void createVariableRow(HSSFWorkbook workbook, HSSFSheet sheet, String xpath, String key, String value, String scope, int formNumber) {
        setVariableCellValue(workbook, sheet, xpath, key, value, scope, getVariablesColumn(formNumber));
    }

    private void setVariableCellValue(HSSFWorkbook workbook, HSSFSheet sheet, String xpath, String key, String value, String scope, int column) {
        final HSSFRow variableRowCell = getVariableRow(workbook, sheet, xpath, key, scope);
        try {
            final double numericalValue = Double.parseDouble(value);
            variableRowCell.createCell(column).setCellValue(DECIMAL_FORMAT.format(numericalValue));
        } catch (NumberFormatException nfe) {
            variableRowCell.createCell(column).setCellValue(value);
        }
        variableRowCell.getCell(column).setCellStyle(getContentStyle(workbook));
    }

    private HSSFRow getVariableRow(HSSFWorkbook workbook, HSSFSheet sheet, String xpath, String key, String scope) {
        if (variableRow.get(xpath + "_" + key) == null) {
            variableRow.put(xpath + "_" + key, sheet.createRow(getVariableRowNumber(xpath, key)));
            variableRow.get(xpath + "_" + key).createCell(VARIABLE_LABEL_COLUMN).setCellValue(key);
            variableRow.get(xpath + "_" + key).getCell(VARIABLE_LABEL_COLUMN).setCellStyle(getAnswerLabelsStyle(workbook));
            variableRow.get(xpath + "_" + key).createCell(VARIABLE_SCOPE_COLUMN).setCellValue(scope);
            variableRow.get(xpath + "_" + key).getCell(VARIABLE_SCOPE_COLUMN).setCellStyle(getAnswerLabelsStyle(workbook));
        }
        return variableRow.get(xpath + "_" + key);
    }

    private int getVariableRowNumber(String xpath, String key) {
        variableRowNumber.computeIfAbsent(xpath + "_" + key, k -> variableRowNumber.size() + TITLE_ROW + 1);
        return variableRowNumber.get(xpath + "_" + key);
    }


    private void createRow(HSSFWorkbook workbook, HSSFSheet sheet, DroolsSubmittedCategory category, DroolsSubmittedQuestion question, int formNumber) {
        setCellValue(workbook, sheet, category, question, getFormResultColumn(formNumber), getAnswersText(question));
    }

    private void setCellValue(HSSFWorkbook workbook, HSSFSheet sheet, DroolsSubmittedCategory category, DroolsSubmittedQuestion question,
                              int column, String value) {
        final HSSFRow rowCell = getQuestionRow(workbook, sheet, category, question);
        try {
            final double numericalValue = Double.parseDouble(value);
            rowCell.createCell(column).setCellValue(DECIMAL_FORMAT.format(numericalValue));
        } catch (NumberFormatException nfe) {
            rowCell.createCell(column).setCellValue(value);
        }
        rowCell.getCell(column).setCellStyle(getContentStyle(workbook));
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

    private int getVariablesColumn(int formNumber) {
        return formNumber + VARIABLE_LABEL_COLUMN;
    }

    private int getFormResultColumn(int formNumber) {
        return formNumber + QUESTION_LABEL_COLUMN;
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

            contentStyle.setAlignment(HorizontalAlignment.CENTER);

            // Border
            contentStyle.setBorderBottom(BorderStyle.THIN);
            contentStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        }
        return contentStyle;
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
