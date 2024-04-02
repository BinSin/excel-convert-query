package com.example.convert.dto;

import io.swagger.v3.oas.annotations.Parameter;
import lombok.Getter;

@Getter
public class UpdateQueryDto extends QueryDto {

  @Parameter(name = "schema", description = "where절 대상 사이즈")
  private int whereSize;

  @Parameter(name = "schema", description = "update 대상 사이즈")
  private int updateSize;

}
