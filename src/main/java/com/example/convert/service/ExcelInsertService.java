package com.example.convert.service;

import com.example.convert.dto.InsertQueryDto;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.context.annotation.Primary;
import org.springframework.stereotype.Service;

@Primary
@Service
@Slf4j
public class ExcelInsertService implements ExcelService<InsertQueryDto> {

  private final int MAX_SIZE = 1000;

  @Override
  public String getQuery(InsertQueryDto queryDto, Sheet worksheet) {
    StringBuilder sb = new StringBuilder();

    String tableName = worksheet.getRow(0).getCell(0).getStringCellValue();
    Row fieldsRow = worksheet.getRow(1); // 두번번째 row 부터 읽기

    int valueLocation = 3;
    int columnSize = fieldsRow.getPhysicalNumberOfCells() - 1;

    String firstQuery = getInsertAndFieldsName(fieldsRow, tableName, columnSize, queryDto.getSchema(), queryDto.getUser());

    // 생성자, 생성일 자동 세팅을 위한 위치 찾기
    int sysCrtrIdIndex = findFieldIndex(fieldsRow, "SYS_CRTR_ID", columnSize);
    int sysCrtDttmIndex = findFieldIndex(fieldsRow, "SYS_CRT_DTTM", columnSize);

    int loopCount = 0;
    while(true) {
      // column 명 세팅
      sb.append(firstQuery);

      // values 세팅
      for (int i = (loopCount * MAX_SIZE) + valueLocation; i < (loopCount+1) * MAX_SIZE + valueLocation; i++) {
        Row row = worksheet.getRow(i);

        if (row == null) { // 첫번째 데이터가 null이면 스톱
          return sb.toString();
        }

        sb.append("(");
        for (int j = 0; j <= columnSize; j++) {
          String value = null;

          Cell cell = row.getCell(j);

          // 값이 null인 경우, null
          if (cell == null) {
            if (j != sysCrtrIdIndex && j != sysCrtDttmIndex) {
              sb.append("null");

              if (j != columnSize) {
                sb.append(", ");
              }
              continue;
            }

            // 생성자, 생성일 자동 세팅
            if (j == sysCrtrIdIndex) {
              value = "ADMIN";
            } else if (j == sysCrtDttmIndex) {
              value = "getdate()";
            }
          } else {
            value = switch (cell.getCellType()) {
              case Cell.CELL_TYPE_STRING -> cell.getStringCellValue();
              case Cell.CELL_TYPE_NUMERIC -> {
                // int형 double형 구별
                double cellValue = cell.getNumericCellValue();
                if (cellValue == Math.rint(cellValue)) {
                  yield String.valueOf((int) cellValue);
                } else {
                  yield String.valueOf(cellValue);
                }
              }
              default -> null;
            };
          }


          // 값이 null인 경우, null
          // 공백인 경우는 공백으로 들어감
          if (value == null) {
            sb.append("null");
          } else {
            sb.append("'").append(value).append("'");
          }

          if (j != columnSize) {
            sb.append(", ");
          }
        }

        if (i != (loopCount+1) * MAX_SIZE + valueLocation - 1 && worksheet.getRow(i+1) != null) {
          sb.append("),\n");
        } else {
          sb.append(");\n");
        }
      }

      loopCount++;
    }
  }

  private String getInsertAndFieldsName(Row fieldsRow, String tableName, int columnSize, String schema, String user) {
    StringBuilder sb = new StringBuilder();

    sb.append("INSERT INTO ").append(schema).append(".").append(user).append(".").append(tableName).append(" (");

    for (int i=0; i <= columnSize; i++) {
      String columnName = fieldsRow.getCell(i).getStringCellValue();

      sb.append(columnName);

      if (i != columnSize) {
        sb.append(", ");
      }
    }
    sb.append(") VALUES\n");

    return sb.toString();
  }

  private int findFieldIndex(Row fieldsRow, String fieldName, int columnSize) {
    for (int i = 0; i <= columnSize; i++) {
      String columnName = fieldsRow.getCell(i).getStringCellValue();
      if (fieldName.equals(columnName)) {
        return i;
      }
    }
    return -1;
  }

}
