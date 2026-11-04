-- 렌즈 품질 검사: 확인한 SQLite DB의 테이블 구조
-- 문서 버전: 2.0 / 확인일: 2026-10-29 (Asia/Seoul)
-- 기준 코드: 50ed1c379945d9b2ab6c0c0233e3e9f2e4cccec3
-- 출처: ../../resources/lens_quality/outputs/results/lens_quality.db
-- 위 경로는 프로젝트 루트(workspace/lens_quality) 기준입니다.
-- DB SHA-256: 4e4c27d6f1c51b79e9205011c8f015491c892664707c0e67a50e225ae8f0e5f6
-- sqlite_master에 저장된 CREATE TABLE 정의를 사용했습니다.
-- 빈 SQLite DB에 실행하면 데이터 없는 4개 테이블(총 47개 컬럼)을 만듭니다.
-- INSERT, PRIMARY KEY, FOREIGN KEY, UNIQUE, NOT NULL, CHECK, DEFAULT,
-- 별도 인덱스는 추가하지 않았습니다. 모든 컬럼은 구조상 NULL을 허용합니다.
-- 이미 같은 테이블이 있으면 CREATE TABLE 오류가 나며, 기존 테이블을 삭제하지 않습니다.
-- Lens.py는 이 파일을 실행하지 않고 DataFrame.to_sql(if_exists="replace")로 저장합니다.
-- to_sql()의 자료형은 입력값에 따라 추론되므로 재생성 시 이 구조와 달라질 수 있습니다.
-- 값·단위·판정 규칙은 database-spec.md, 처리 순서는 data-flow.md를 참고하세요.

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
