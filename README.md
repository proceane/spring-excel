# spring-excel

주요 기능
1. 엑셀 생성
2. Rest API를 사용하여 다운로드
3. 암호화된 엑셀 파일 생성 및 다운로드

### Environment
- Java : 8
- Spring Boot : 2.7.8
- POI: 3.16

### API
1. 엑셀 파일 생성 및 다운로드
```text
http://localhost:8080/excel
```
2. 암호화 엑셀 파일 생성 및 다운로드(password: 1234)
```text
http://localhost:8080/excel/xssf/encrypt
http://localhost:8080/excel/sxssf/encrypt
```