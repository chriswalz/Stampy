package stampy;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Stampy {
    private Pattern p = Pattern.compile("\\{\\{(.+)\\}\\}"); // matches pattern {{...}}
    private SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
    private String path;
    private Workbook wb;

    public static Stampy openTemplate(String path) {
        return new Stampy(path);
    }

    private Stampy(String path) {
        this.path = path;
        try {
            InputStream fileStream = new FileInputStream(path);
            this.wb = WorkbookFactory.create(fileStream);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
            System.exit(-1);
        }
    }

    public void printTemplateNames() {
        List<? extends Name> names = wb.getAllNames();
        System.out.println("Number of names: " + names);
    }
    public void executeTemplateMustaches(String output, Map<String, Object> ctxFieldMap) {
        int fail = 0;
        Sheet sheet;
        for (int a = 0; a < wb.getNumberOfSheets(); a++) {
            sheet = wb.getSheetAt(a);
            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                Row r = sheet.getRow(i);
                if (r == null) {
                    continue;
                }
                for (int j = 0; j < r.getLastCellNum(); j++) {
                    Cell c = r.getCell(j);
                    if (c == null) {
                        continue;
                    }
                    if (c.getCellTypeEnum() == CellType.STRING) {
                        Matcher m = p.matcher(c.getStringCellValue());
                        if (m.matches()) {
                            String rpField = m.group(1);
                            Object val = ctxFieldMap.get(rpField);
                            fail += setCellValue(c, val);
                        }
                    }
                }
            }
        }
        if (fail < 0) {
            throw new RuntimeException("Not enough map values supplied");
        }
        // calculate formulas
        wb.setForceFormulaRecalculation(true);
        try {
            OutputStream fileOut = new FileOutputStream(output);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
            System.exit(-1);
        }
    }

    private int setCellValue(Cell cell, Object ctx) {
        //CellStyle style = cell.getCellStyle();
        if (ctx == null) {
            /*cell.setCellValue("{{error}}");*/
            System.err.println("Cell not set: " + cell.getStringCellValue());
            return -1;
        } else if (ctx instanceof Integer) {
            cell.setCellValue(((Integer) ctx));
        } else if (ctx instanceof String) {
            cell.setCellValue(((String) ctx));
        } else if (ctx instanceof LocalDate) {
            cell.setCellValue(ctx.toString());
        } else if (ctx instanceof Double) {
            cell.setCellValue(((Double) ctx));
        } else if (ctx instanceof Long) {
            cell.setCellValue(((Long) ctx));
        } else if (ctx instanceof Date) {
            cell.setCellValue(dateFormat.format((Date) ctx));
        } else if (ctx.getClass().isArray()) {
            // insert row instead?
            int rI = cell.getRow().getRowNum() + 1;
            Object[] arr = (Object[]) ctx;
            cell.getSheet().shiftRows(rI, cell.getSheet().getLastRowNum(), arr.length - 1);
            int io = cell.getColumnIndex() + 1;
            for (Object obj : arr) {
                int i = cell.getColumnIndex() + 1;
                if (obj.getClass().isArray()) {
                    Object[] innerArr = (Object[]) obj;
                    for (Object inner : innerArr) {
                        Cell tmp = getCell(cell.getSheet(), i++, rI);
                        //tmp.setCellStyle(style);
                        setCellValue(tmp, inner);
                    }
                    rI++;
                } else {
                    Cell tmp = getCell(cell.getSheet(), io++, rI);
                    //tmp.setCellStyle(style);
                    setCellValue(tmp, obj);
                }
            }
        } else {
            throw new RuntimeException(ctx.toString());
        }
        return 0;
    }

    private Cell getCell(Sheet sheet, int x, int y) {
        y--;
        x--;
        Row row = sheet.getRow(y);
        if (row == null) {
            row = sheet.createRow(y);
        }
        // copy formatting
        Cell cell = row.getCell(x);
        if (cell == null) {
            cell = row.createCell(x);
        }
        return cell;
    }

    public void setDateFormat(SimpleDateFormat dateFormat) {
        this.dateFormat = dateFormat;
    }


    /*
      stampy.executeTemplateMustaches(
                "stampy_output.xls",
                new Object() {
                    public double rate = 10.4;
                    public int profit = 100_100;
                }
        );
    public void executeTemplateMustaches(String output, Object ctx) {
        Map<String, Object> ctxFieldMap = new HashMap<>();
        for (Field field : ctx.getClass().getDeclaredFields()) {
            if (field.getName().contains("this$0")) {
                continue;
            }
            try {
                ctxFieldMap.put(field.getName(), field.get(ctx));
            } catch (IllegalArgumentException | IllegalAccessException e) {
                ctxFieldMap.put(field.getName(), "{{error}}");
                e.printStackTrace();
            }
        }

        executeTemplateMustaches(output, ctxFieldMap);
    } */
}
