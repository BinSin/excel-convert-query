package com.example.convert.service;

import com.example.convert.dto.UpdateQueryDto;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

@Service
@Slf4j
public class ExcelUpdateService implements ExcelService<UpdateQueryDto> {

  private final int MAX_SIZE = 1000;
  private final String SCHEMA = "FCM";
  private final String USER = "dbo";

  @Override
  public String getQuery(UpdateQueryDto queryDto, Sheet worksheet) {
    return "";
  }

}
