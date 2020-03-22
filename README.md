# crawl_applicant
파이썬을 활용하여, 구글 폼으로 받은 지원서 중 면접 대상자 지원서만 크롤링하기

## stack
python 3.8.0

et-xmlfile==1.0.1
jdcal==1.4.1
openpyxl==3.0.3



## 사용안내
data 폴더 안에 아래 2가지 데이터를 준비합니다.
1. application.csv: 구글폼으로 받은 지원자들의 서류를 구글 폼 내에서 csv로 다운받아, application.csv라는 이름으로 수정합니다.
2. timetable.xlsx: 구글 공유 문서에서 배포용이 아닌 면접 타임테이블을 우클릭하여 xlsx 파일로 다운받아, timetable.xlsx라는 이름으로 수정합니다.   
   
이후 해당 repo의 make.py 파일을 실행하여, data 폴더 안에 생성되는 online_data, offline_data 파일을 확인한다.   
   

> ### Repo 안내
> 해당 repository내의 모든 내용에 대한 권한은 문범우에게 있습니다.   
> 관련 내용에 대한 무단도용 및 무단참조는 허가하지 않습니다.   
> E: doorbw@outlook.com