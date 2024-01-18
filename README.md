<div><img width="418" alt="image" src="https://github.com/20161609/FinanceTreeProject/assets/92266688/d6a43b42-ae2c-4db7-8673-470ae91c1c99"></div>
<br>
<br>

# 🌳 프로젝트 소개 (Project Description)
Finance Tree는 금융 거래를 관리하고 조직화하는 커맨드 라인 인터페이스 도구입니다. 사용자는 마치 파일 시스템을 탐색하듯, 금융 거래를 트리 구조로 관리할 수 있습니다. 이 도구는 SQLite 데이터베이스를 사용하여 거래 내역을 저장하며, 다양한 커맨드를 통해 데이터를 쉽게 관리할 수 있습니다.

# 📆 개발 기간 (Period)
2024-01-10 ~ 2024-01-18

# ⚙️ 기능 (Features)
- 트리 구조 관리: 금융 거래를 트리 구조로 관리하여 직관적인 조직화 가능
- 다양한 커맨드 지원: mkdir, rmdir, cd, ls 등의 파일 시스템 커맨드와 유사한 작업 수행
- 데이터베이스 연동: SQLite를 이용한 거래 데이터 저장 및 관리
- 사용자 친화적 인터페이스: 커맨드 라인을 통한 간편한 사용자 인터랙션

# 🔧 설치 방법 (Installation)
이 프로젝트를 사용하기 위해서는 다음 단계를 따르세요.

1. **리포지토리 복제(Clone the repository):**
   ```bash
   https://github.com/20161609/FinanceTreeProject.git

2. **압축 해제**
   executable.zip을 원하는 위치에 압축 해제합니다.



# 📚 사용 방법 (Usage)
*Commands: 기존의 쉘명령어와 유사합니다.
<br>

**mkdir, md <dir_name> : 새 디렉토리를 생성합니다.**
<div><img width="418" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/00aaf632-3a18-4086-bb24-e6ecdc9e04dc"></div>
<br>
<br>

**rmdir, rd <dir_name> : 기존 디렉토리를 삭제합니다.**
**rmdir, rd <dir_number> : 인덱스 번호를 기반으로 한 기존 디렉토리를 삭제합니다.**
<div><img width="534" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/2303035a-7440-45c5-a296-772366f702ed"></div>
<br>
<br>

**chdir, cd <path> : 현재 디렉토리를 변경합니다.**
**chdir, cd <dir_number> : 인덱스 번호를 기반으로 현재 디렉토리를 변경합니다.**
**list, ls : 현재 디렉토리의 모든 디렉토리를 나열합니다.**
<div><img width="367" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/d4ddcf99-8ec8-47b8-8c50-8f7c889d56ad"></div>
<br>
<br>

**refer, rf -d <period> : 일일 거래를 참조합니다. 기간을 선택적으로 지정할 수 있습니다 (예: 2023/01/01~2023/01/31).**
<div><img width="721" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/f526daa9-2678-4dc2-8b08-297733e90407">   </div>
<div><img width="688" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/126a3d46-bdc6-4216-ac51-f06ac426ce64"></div>
<br>
<br>
      
**refer, rf -m <period> : 월별 거래를 참조합니다. 기간을 선택적으로 지정할 수 있습니다 (예: 2023/01/01~2023/01/31).**
<div><img width="366" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/74826492-fc9b-4340-95ba-949e44ac0b2f"></div>
<br>
<br>

**refer, rf -t : 거래의 트리 구조를 표시합니다.**
<div><img width="465" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/86519df0-696d-4877-9fc5-1af0e650f77b"></div>
<br>
<br>

**insert : 새로운 거래 기록을 삽입합니다.**
<div><img width="945" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/fb7f98b4-60ed-4bfc-a026-907e905faade">   </div>   
<br>
<br>

**delete : 기존의 거래 기록을 삭제합니다. (업데이트 날짜로 정렬된 테이블)**
**delete -t : 기존의 거래 기록을 삭제합니다. (거래 날짜로 정렬된 테이블)**

<div><img width="741" alt="image" src="https://github.com/20161609/FinanceTree/assets/92266688/975da600-994f-459c-bea6-7d187dcd7df3"></div>
<br>
<br>
   
**help : 이 도움말 메시지를 표시합니다.**
<div><img width="830" alt="image" src="https://github.com/20161609/FinanceTreeProject/assets/92266688/68a791ec-5ef9-4a41-a054-040d3c0e8a30"></div>
<br>
<br>

<div>* 어떤 작업을 취소하거나 종료하고 싶을 때 언제든지 'q!'를 사용하세요.</div>
<div>
* 날짜 입력은 다양한 입력 방식으로 가능합니다. (예시, 2024년 1월 18일의 경우)
   <div>- "2024-01-18", "2024/01/18"</div>
   <div>- "20240118", "240018" (2000년대 일 경우만 가능) </div>
   <div>- "24-01-18"(2000년대 일 경우만 가능), "24/01/18"(2000년대 일 경우만 가능)</div>
</div>

# 🙋 저자: 안호준 (dksshwns@sogang.ac.kr)
