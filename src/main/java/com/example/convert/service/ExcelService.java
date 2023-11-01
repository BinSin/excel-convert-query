package com.example.convert.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
@Slf4j
public class ExcelService {

  private final int MAX_SIZE = 1000;
  private final String SCHEMA = "FCM";
  private final String USER = "dbo";

  public String convertToQuery(MultipartFile file) throws Exception {
    String extension = file.getOriginalFilename().split("\\.")[1];

    Workbook workbook = null;
    if (extension.equals("xlsx")) {
      workbook = new XSSFWorkbook(file.getInputStream());
    } else if (extension.equals("xls")) {
      workbook = new HSSFWorkbook(file.getInputStream());
    }

    StringBuilder sb = new StringBuilder();
    int sheetSize = workbook.getNumberOfSheets();
    for (int i=0; i<sheetSize; i++) {
      Sheet worksheet = workbook.getSheetAt(i);

      sb.append(covertRowToInsertQuery(worksheet)).append("\n");
    }

    return sb.toString();
  }

  private String covertRowToInsertQuery(Sheet worksheet) {
    StringBuilder sb = new StringBuilder();

    String tableName = worksheet.getSheetName();
    Row columnRow = worksheet.getRow(0);
    int columnSize = columnRow.getPhysicalNumberOfCells() - 1;

    int s = 0;
    while(true) {
      // column 명 세팅
      sb.append("INSERT INTO ").append(SCHEMA).append(".").append(USER).append(".").append(tableName).append(" (");

      // 첫번째 row 읽기
      for (int i=0; i<= columnSize; i++) {
        String columnName = columnRow.getCell(i).getStringCellValue();
        sb.append(columnName);

        if (i != columnSize) {
          sb.append(", ");
        }
      }
      sb.append(") VALUES\n");

      // values 세팅
      for (int i = (s * MAX_SIZE) + 1; i <= (s+1) * MAX_SIZE; i++) {
        Row row = worksheet.getRow(i);

        if (row == null) { // 첫번째 데이터가 null이면 스톱
          return sb.toString();
        }

        sb.append("(");
        for (int j = 0; j <= columnSize; j++) {

          Cell cell = row.getCell(j);
          // // 값이 null인 경우, null
          if (cell == null) {
            sb.append("null");

            if (j != columnSize) {
              sb.append(", ");
            }
            continue;
          }

          String value = switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING -> cell.getStringCellValue();
            case Cell.CELL_TYPE_NUMERIC -> Integer.toString((int) cell.getNumericCellValue());
            default -> null;
          };

          // 값이 null인 경우, null
          // 공백인 경우는 공백으로 들어감
          if (value == null) {
            sb.append("null");
          } else {
            if (value.equals("getdate()")) {
              sb.append(value);
            } else {
              sb.append("'").append(value).append("'");
            }
          }

          if (j != columnSize) {
            sb.append(", ");
          }
        }

        if (i != (s+1) * MAX_SIZE && worksheet.getRow(i+1) != null) {
          sb.append("),\n");
        } else {
          sb.append(");\n");
        }
      }

      s++;
    }
  }

  private String covertRowToUpdateQuery(Sheet worksheet) {
    StringBuilder sb = new StringBuilder();

    String tableName = worksheet.getRow(1).getCell(0).getStringCellValue();
    sb.append("UPDATE FCM.dbo.").append(tableName).append(" SET ").append("");
    log.info("table Name: {}", tableName);

    for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
      Row row = worksheet.getRow(i);


    }

    return sb.toString();
  }

}
