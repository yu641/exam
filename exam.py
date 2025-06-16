import os
import time
import shutil
import logging
import pandas as pd
import win32com.client as win32
import win32clipboard as wc
import win32con

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 한글 열기
def open_hwp():
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    return hwp

# 클립보드로 텍스트 삽입
def insert_tag_via_clipboard(hwp, text):
    wc.OpenClipboard()
    wc.EmptyClipboard()
    wc.SetClipboardData(win32con.CF_UNICODETEXT, text)
    wc.CloseClipboard()
    hwp.HAction.Run("MoveDocEnd")
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("BreakPara")

# 서식 유지한 .hwp 복사 붙여넣기
def insert_formatted_content(hwp, src_path):
    if not os.path.exists(src_path):
        logger.warning(f"파일 없음: {src_path}")
        return False

    temp = open_hwp()
    temp.Open(src_path, "", "")
    time.sleep(0.5)

    temp.HAction.Run("SelectAll")
    temp.HAction.Run("Copy")
    time.sleep(0.5)

    hwp.HAction.Run("MoveDocEnd")
    hwp.HAction.Run("Paste")
    time.sleep(0.5)

    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("BreakPara")

    temp.Quit()
    return True

# 시험지 문서 생성
def create_exam_doc(output_path, problem_paths):
    template_path = os.path.join(os.getcwd(), "빈_템플릿_2단.hwp")
    if not os.path.exists(template_path):
        logger.error("빈_템플릿_2단.hwp 파일이 없습니다.")
        return False

    try:
        shutil.copy(template_path, output_path)
    except Exception as e:
        logger.error(f"파일 복사 실패: {e}")
        return False

    hwp = None
    try:
        hwp = open_hwp()
        hwp.Open(output_path,"","")
        time.sleep(2.0)

        hwp.XHwpWindows.Item(0).Visible = True
        time.sleep(0.5)
        hwp.HAction.Run("MoveDocBegin")

        passage_number = 1
        problem_number_in_passage = 1
        current_passage_id = None

        for i, (tag, path) in enumerate(problem_paths):
            logger.info(f"{i+1}번 삽입 중: {tag}")

            if tag.startswith("지문 "):
                insert_tag_via_clipboard(hwp, f"{passage_number}.")
                insert_formatted_content(hwp, path)
                current_passage_id = tag.split()[1]
                problem_number_in_passage = 1
                passage_number += 1

            elif tag.startswith("문제 "):
                this_pid = tag.split()[1].rsplit('_', 1)[0]
                if this_pid != current_passage_id:
                    logger.warning(f"문제가 이전 지문과 매칭되지 않음: {tag}")  # 다른 문제가 동일 지문 참조시 2번 출력되지 않게 함
                insert_tag_via_clipboard(hwp, f"{passage_number-1}-{problem_number_in_passage})")
                insert_formatted_content(hwp, path)
                problem_number_in_passage += 1

        hwp.Save()
        logger.info(f"파일 저장 완료")
        return True
    except Exception as e:
        logger.error(f"문서 생성 오류: {e}")
        return False
    finally:
        if hwp:
            try:
                hwp.Quit()
            except:
                pass

# 정답률 → 난이도
def classify_difficulty(rate):
    if pd.isna(rate):
        return None
    elif rate <= 0.60:
        return "상"
    elif rate <= 0.80:
        return "중"
    else:
        return "하"

# 시험지 생성
def generate_exam_sheet(excel_path, base_dir, subject="", passage_type="", level="", num_questions=5):
    logger.info("시험지 생성 시작")
    try:
        df = pd.read_excel(excel_path)
        logger.info(f"Excel 읽기 완료: {len(df)}행")
    except Exception as e:
        logger.error(f"Excel 읽기 실패: {e}")
        return

    df["난이도"] = df["정답률"].apply(classify_difficulty)
    df_problems = df[df["유형"] == "문제"]

    # 조건이 빈 문자열이면 필터 제외
    conditions = (df_problems["유형"] == "문제")
    if subject and subject.strip():
        conditions &= (df_problems["과목"] == subject)
    if passage_type and passage_type.strip():
        conditions &= (df_problems["지문유형"] == passage_type)
    if level and level.strip():
        conditions &= (df_problems["난이도"] == level)

    filtered = df_problems[conditions]

    if filtered.empty:
        logger.warning("조건에 맞는 문제가 없습니다.")
        print("조건에 맞는 문제가 없습니다.")
        return

    selected = filtered.sample(n=min(num_questions, len(filtered)), random_state=42)
    problem_paths = []
    added_passages = set()

    for _, row in selected.iterrows():
        qid = row["문제id"]
        pid = row["지문id"]

        prob_path = os.path.join(base_dir, "문제", f"{qid}.hwp")
        passage_path = os.path.join(base_dir, "지문", f"{pid}.hwp")

        if os.path.exists(passage_path) and pid not in added_passages:
            problem_paths.append((f"지문 {pid}", passage_path))
            added_passages.add(pid)

        if os.path.exists(prob_path):
            problem_paths.append((f"문제 {qid}", prob_path))
        else:
            logger.warning(f"문제 파일 없음: {prob_path}")

    if not problem_paths:
        logger.warning("추출된 파일이 없습니다.")
        print("추출된 파일이 없습니다.")
        return

    # 출력 파일 이름 구성
    label = f"{subject or '전체'}_{passage_type or '전체'}_{level or '전체'}"
    output_name = f"{label}.hwp"
    output_path = os.path.join(base_dir, output_name)

    if create_exam_doc(output_path, problem_paths):
        print(f"시험지가 생성되었습니다: {output_name}")
    else:
        print("시험지 생성에 실패했습니다.")
