-- 렌즈 검사 프로젝트: 현재 SQLite DB의 테이블 구조
-- 확인일: 2026-10-24
-- 출처: ../../resources/lens_quality/outputs/results/lens_quality.db
-- 위 경로는 프로젝트 루트(lens_quality)를 기준으로 합니다.
-- sqlite_master에 저장된 CREATE TABLE 문을 기준으로 작성했습니다.
-- 데이터(INSERT), 기본키, 외래키, NOT NULL, CHECK, 별도 인덱스는 추가하지 않았습니다.
-- 빈 SQLite DB에 실행하면 아래 네 개의 빈 테이블을 만듭니다.
-- 이미 같은 테이블이 있으면 오류가 나며, 기존 테이블을 삭제하거나 덮어쓰지 않습니다.
-- Lens.py는 이 파일을 실행하지 않습니다. 실제 테이블은 DataFrame.to_sql()로 만듭니다.
-- 원본 데이터의 자료형이 바뀌면 to_sql()이 추론하는 DB 자료형도 달라질 수 있습니다.

CREATE TABLE "measurements" (
  "제품번호" TEXT,
  "소재종류" TEXT,
  "제품명" TEXT,
  "스크래치등급" INTEGER,
  "기포이물질개수" INTEGER,
  "코팅불량개수" INTEGER,
  "굴절률" REAL,
  "광투과율" REAL,
  "왜곡률" REAL,
  "렌즈직경" REAL,
  "중심두께" REAL,
  "가장자리두께" REAL,
  "곡률반경" REAL,
  "UV차단율" REAL,
  "색상균일도" REAL,
  "틴트농도" REAL
);

CREATE TABLE "criteria" (
  "검사항목" TEXT,
  "적용소재" TEXT,
  "단위" TEXT,
  "하한" REAL,
  "목표치" REAL,
  "상한" REAL
);

CREATE TABLE "inspection_results" (
  "제품번호" TEXT,
  "소재종류" TEXT,
  "제품명" TEXT,
  "스크래치등급" INTEGER,
  "기포이물질개수" INTEGER,
  "코팅불량개수" INTEGER,
  "굴절률" REAL,
  "광투과율" REAL,
  "왜곡률" REAL,
  "렌즈직경" REAL,
  "중심두께" REAL,
  "가장자리두께" REAL,
  "곡률반경" REAL,
  "UV차단율" REAL,
  "색상균일도" REAL,
  "틴트농도" REAL,
  "판정" TEXT,
  "판정사유" TEXT
);

CREATE TABLE "inspection_details" (
  "제품번호" TEXT,
  "검사항목" TEXT,
  "측정값" REAL,
  "하한" REAL,
  "상한" REAL,
  "항목판정" TEXT,
  "사유" TEXT
);
