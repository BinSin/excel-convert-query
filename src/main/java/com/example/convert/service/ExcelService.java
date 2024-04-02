package com.example.convert.service;

import com.example.convert.dto.InsertHistoryQueryDto;
import com.example.convert.dto.InsertQueryDto;
import com.example.convert.dto.UpdateQueryDto;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelService {

  private final int MAX_SIZE = 1000;

  public String convertInsertQuery(InsertQueryDto queryDto, MultipartFile file) throws Exception {
    Sheet worksheet = getWorkbook(file).getSheetAt(0);

    StringBuilder sb = new StringBuilder();

    Row fieldsRow = worksheet.getRow(0); // 두번번째 row 부터 읽기

    int valueLocation = 1;
    int columnSize = fieldsRow.getPhysicalNumberOfCells() - 1;

    String firstQuery = getInsertAndFieldsName(fieldsRow, queryDto.getTableName(), columnSize,
        queryDto.getSchema(), queryDto.getUser());

    // 생성자, 생성일 자동 세팅을 위한 위치 찾기
    int sysCrtrIdIndex = findFieldIndex(fieldsRow, "SYS_CRTR_ID", columnSize);
    int sysCrtDttmIndex = findFieldIndex(fieldsRow, "SYS_CRT_DTTM", columnSize);

    int loopCount = 0;
    while (true) {
      // column 명 세팅
      sb.append(firstQuery);

      // values 세팅
      for (int i = (loopCount * MAX_SIZE) + valueLocation;
          i < (loopCount + 1) * MAX_SIZE + valueLocation; i++) {
        Row row = worksheet.getRow(i);

        if (row == null) { // 첫번째 데이터가 null이면 스톱
          return sb.toString();
        }

        sb.append("(");
        for (int j = 0; j <= columnSize; j++) {
          String value = null;

          Cell cell = row.getCell(j);

          if (j == sysCrtrIdIndex) {
            value = "ADMIN";
          } else if (j == sysCrtDttmIndex) {
            value = "getdate()";
          } else if (cell == null) { // 값이 null인 경우, null
            if (j != sysCrtrIdIndex && j != sysCrtDttmIndex) {
              sb.append("null");

              if (j != columnSize) {
                sb.append(", ");
              }
              continue;
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
          } else if (j == sysCrtDttmIndex) { // 생성일은 value 그대로 입력
            sb.append(value);
          } else {
            sb.append("'").append(value).append("'");
          }

          if (j != columnSize) {
            sb.append(", ");
          }
        }

        if (i != (loopCount + 1) * MAX_SIZE + valueLocation - 1
            && worksheet.getRow(i + 1) != null) {
          sb.append("),\n");
        } else {
          sb.append(");\n");
        }
      }

      loopCount++;
    }
  }

  private String getInsertAndFieldsName(Row fieldsRow, String tableName, int columnSize,
      String schema, String user) {
    StringBuilder sb = new StringBuilder();

    sb.append("INSERT INTO ").append(schema).append(".").append(user).append(".").append(tableName)
        .append(" (");

    for (int i = 0; i <= columnSize; i++) {
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

  public String convertUpdateQuery(UpdateQueryDto queryDto, MultipartFile file) throws Exception {
    StringBuilder sb = new StringBuilder();

    Sheet worksheet = getWorkbook(file).getSheetAt(0);

    Row columNameRow = worksheet.getRow(0);
    int columnSize = columNameRow.getPhysicalNumberOfCells() - 1;

    int updateSize = queryDto.getUpdateSize();
    int whereSize = queryDto.getWhereSize();

    int sysCrtrIdIndex = findFieldIndex(columNameRow, "SYS_UPDR_ID", columnSize);
    int sysCrtDttmIndex = findFieldIndex(columNameRow, "SYS_UPDT_DTTM", columnSize);

    int index = 1;
    while (true) {
      Row row = worksheet.getRow(index);

      if (row == null) { // 첫번째 데이터가 null이면 스톱
        return sb.toString();
      }

      sb.append("UPDATE ").append(queryDto.getSchema())
          .append(".").append(queryDto.getUser()).append(".").append(queryDto.getTableName())
          .append("\n")
          .append("SET ");

      int endIndex = whereSize + updateSize;
      for (int i = whereSize; i < endIndex; i++) {
        Cell nameCell = columNameRow.getCell(i);
        sb.append(nameCell.getStringCellValue()).append(" = ");

        String value;
        if (i == sysCrtrIdIndex) {
          value = "ADMIN";
        } else if (i == sysCrtDttmIndex) {
          value = "getdate()";
        } else {
          Cell valueCell = row.getCell(i);
          value = switch (valueCell.getCellType()) {
            case Cell.CELL_TYPE_STRING -> valueCell.getStringCellValue();
            case Cell.CELL_TYPE_NUMERIC -> {
              // int형 double형 구별
              double cellValue = valueCell.getNumericCellValue();
              if (cellValue == Math.rint(cellValue)) {
                yield String.valueOf((int) cellValue);
              } else {
                yield String.valueOf(cellValue);
              }
            }
            default -> null;
          };

        }

        if (i == sysCrtDttmIndex) {
          sb.append(value);
        } else {
          sb.append("'").append(value).append("'");
        }

        if (i != endIndex - 1) {
          sb.append(", ");
        }
      }

      sb.append("\nWHERE ");

      for (int i = 0; i < whereSize; i++) {
        Cell nameCell = columNameRow.getCell(i);
        sb.append(nameCell.getStringCellValue()).append(" = ");

        Cell valueCell = row.getCell(i);
        String value = switch (valueCell.getCellType()) {
          case Cell.CELL_TYPE_STRING -> valueCell.getStringCellValue();
          case Cell.CELL_TYPE_NUMERIC -> {
            // int형 double형 구별
            double cellValue = valueCell.getNumericCellValue();
            if (cellValue == Math.rint(cellValue)) {
              yield String.valueOf((int) cellValue);
            } else {
              yield String.valueOf(cellValue);
            }
          }
          default -> null;
        };

        if (value == null) {
          sb.append("null");
        } else {
          sb.append("'").append(value).append("'");
        }

        if (i != whereSize - 1) {
          sb.append(" and ");
        }
      }

      sb.append(";\n");

      index++;
    }
  }

  private Workbook getWorkbook(MultipartFile file) throws Exception {
    String extension = file.getOriginalFilename().split("\\.")[1];
    return switch (extension) {
      case "xlsx" -> new XSSFWorkbook(file.getInputStream());
      case "xls" -> new HSSFWorkbook(file.getInputStream());
      default -> throw new Exception("잚못된 파일 형식입니다.");
    };
  }

  public String convertInsertHistoryQuery(InsertHistoryQueryDto queryDto, MultipartFile file)
      throws Exception {
    Sheet worksheet = getWorkbook(file).getSheetAt(0);

    StringBuilder sb = new StringBuilder();

    Row fieldsRow = worksheet.getRow(0); // 두번번째 row 부터 읽기

    int valueLocation = 1;
    int columnSize = fieldsRow.getPhysicalNumberOfCells() - 1;

    String firstQuery = getInsertAndFieldsName(fieldsRow, queryDto.getTableName(), columnSize,
        queryDto.getSchema(), queryDto.getUser());

    int infoUpdtDttmIndex = findFieldIndex(fieldsRow, "INFO_UPDT_DTTM", columnSize);

    int loopCount = 0;
    while (true) {
      // column 명 세팅
      sb.append(firstQuery);

      // values 세팅
      for (int i = (loopCount * MAX_SIZE) + valueLocation;
          i < (loopCount + 1) * MAX_SIZE + valueLocation; i++) {
        Row row = worksheet.getRow(i);

        if (row == null) { // 첫번째 데이터가 null이면 스톱
          return sb.toString();
        }

        sb.append("(");
        for (int j = 0; j <= columnSize; j++) {
          String value = null;

          Cell cell = row.getCell(j);

          if (j == infoUpdtDttmIndex) {
            value = "getdate()";
          } else if (cell == null) { // 값이 null인 경우, null
            if (j != infoUpdtDttmIndex) {
              sb.append("null");

              if (j != columnSize) {
                sb.append(", ");
              }
              continue;
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
            if (j == infoUpdtDttmIndex) {
              sb.append(value);
            } else {
              sb.append("'").append(value).append("'");
            }
          }

          if (j != columnSize) {
            sb.append(", ");
          }
        }

        if (i != (loopCount + 1) * MAX_SIZE + valueLocation - 1
            && worksheet.getRow(i + 1) != null) {
          sb.append("),\n");
        } else {
          sb.append(");\n");
        }
      }

      loopCount++;
    }
  }

}
