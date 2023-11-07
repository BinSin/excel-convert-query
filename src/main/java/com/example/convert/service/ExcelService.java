package com.example.convert.service;

import com.example.convert.dto.QueryDto;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

public interface ExcelService<T extends QueryDto> {

  default String convertToQuery(T queryDto, MultipartFile file) throws Exception {
    Workbook workbook = getWorkbook(file);
    return getQuery(queryDto, workbook.getSheetAt(0)); // 첫번째 시트만 처리
  }

  default Workbook getWorkbook(MultipartFile file) throws Exception {
    String extension = file.getOriginalFilename().split("\\.")[1];
    return switch (extension) {
      case "xlsx" -> new XSSFWorkbook(file.getInputStream());
      case "xls" -> new HSSFWorkbook(file.getInputStream());
      default -> throw new Exception("잚못된 파일 형식입니다.");
    };
  }

  String getQuery(T queryDto, Sheet workbook);

}
