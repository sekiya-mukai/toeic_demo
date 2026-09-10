from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
import logging


BASE_DIR = Path(__file__).parent.parent

question_filename = "問題用紙.xlsx"
answer_filename = "解答用紙.xlsx"

QUESTION_FILE = BASE_DIR / question_filename
ANSWER_FILE = BASE_DIR / answer_filename
LOG_FILE = BASE_DIR / "output.log"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8"
)

def log_print(message=""):
    print(message)
    logging.info(message)


def load_answers(ws):
    """
    A2:J21 を読み込み、
    {問題番号: 解答}
    の辞書を返す
    """

    answers = {}

    for row in range(2, 22, 2):

        question_row = row
        answer_row = row + 1

        for col in range(1, 11):

            question_no = ws.cell(question_row, col).value
            answer = ws.cell(answer_row, col).value

            if question_no is None:
                continue

            answers[int(question_no)] = str(answer).strip().upper()

    return answers


def main():

    log_print("=== TOEIC Demo ===")

    selected_type = input(
        "L (Listening) または R (Reading) を入力: "
    ).strip().upper()
    log_print(f"[DEBUG]: selected_type.upper(): {selected_type.upper()}")
    if selected_type.upper() not in ["L", "R"]:
        log_print("L または R を入力してください")
        return

    # 受験タイプと番号を決定(リスニング or リーディング, テスト番号(1~5))
    selected_test_no = input(
        "問題番号 (1～5) を入力: "
    ).strip()
    if selected_test_no not in ["1", "2", "3", "4", "5"]:
        log_print("問題番号は 1～5 を入力してください")
        return
    answer_sheetname = f"{selected_type}_{selected_test_no}"
    log_print(f"[DEBUG]: answer_sheetname: {answer_sheetname}")

    # 問題用紙読み込み
    question_wb = load_workbook(QUESTION_FILE)
    question_sheetname = "回答用紙"
    if question_sheetname not in question_wb.sheetnames:
        log_print(f"{question_filename} に [{question_sheetname}] シートが存在しません")
        return
    # 受験者回答シート
    user_sheet = question_wb[question_sheetname]
    user_answers = load_answers(user_sheet)


    # 解答用紙読み込み
    answer_wb = load_workbook(ANSWER_FILE)
    if answer_sheetname not in answer_wb.sheetnames:
        log_print(
            f"{answer_filename} に "
            f"[{answer_sheetname}] シートが存在しません"
        )
        return

    # 正解シート（固定）
    correct_sheet = answer_wb[answer_sheetname]
    correct_answers = load_answers(correct_sheet)
    total = len(correct_answers)
    correct = 0
    wrong_questions = []

    for question_no, correct_answer in correct_answers.items():
        user_answer = user_answers.get(question_no)
        if user_answer == correct_answer:
            correct += 1
        else:
            wrong_questions.append(question_no)

    wrong = total - correct

    accuracy = round(
        correct / total * 100,
        2
    )

    if wrong_questions:
        log_print()
        log_print("不正解問題:")
        log_print(", ".join(map(str, wrong_questions)))

    log_print()
    log_print("=== 採点結果 ===")
    log_print(f"正答数 : {correct}")
    log_print(f"誤答数 : {wrong}")
    log_print(f"正答率 : {accuracy}%")


    # 結果シート
    result_sheet_name = "結果"

    if result_sheet_name not in answer_wb.sheetnames:

        result_ws = answer_wb.create_sheet(result_sheet_name)

        result_ws["A1"] = "実施日"
        result_ws["B1"] = "Listening"
        result_ws["C1"] = "Reading"
        result_ws["D1"] = "正答"
        result_ws["E1"] = "正答率"

    else:
        result_ws = answer_wb[result_sheet_name]

    row = result_ws.max_row + 1

    result_ws[f"A{row}"] = datetime.now().strftime("%Y/%m/%d")

    if selected_type == "L":
        result_ws[f"B{row}"] = selected_test_no
        result_ws[f"C{row}"] = ""
    else:
        result_ws[f"B{row}"] = ""
        result_ws[f"C{row}"] = selected_test_no

    result_ws[f"D{row}"] = correct
    result_ws[f"E{row}"] = accuracy

    answer_wb.save(ANSWER_FILE)

    log_print()
    log_print("結果シートへ保存しました")


if __name__ == "__main__":
    main()