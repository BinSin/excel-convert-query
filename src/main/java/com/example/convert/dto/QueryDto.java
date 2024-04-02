package com.example.convert.dto;

import io.swagger.v3.oas.annotations.Parameter;
import lombok.Getter;

@Getter
public class QueryDto {

  @Parameter(name = "schema", description = "DB 스키마명")
  private String schema;

  @Parameter(name = "user", description = "DB 사용자명")
  private String user;

  @Parameter(name = "schema", description = "DB 테이블명")
  private String tableName;

}
