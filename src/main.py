from openpyxl import load_workbook
from datetime import datetime


QUESTION_FILE = "問題用紙.xlsx"
RESULT_FILE = "結果用紙.xlsx"


def load_answers(ws):
    """
    A2:J21 の解答シートを読み込み
    {問題番号: 解答}
    の辞書を返却
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

    print("=== TOEIC採点システム ===")

    # 2. Listening or Reading選択
    selected_type = input(
        "Listening または Reading を入力してください: "
    ).strip()

    if selected_type not in ["Listening", "Reading"]:
        print("入力エラー")
        return

    # 3. 問題番号選択
    try:
        selected_test_no = int(
            input("問題番号(1～5)を入力してください: ")
        )
    except ValueError:
        print("入力エラー")
        return

    if selected_test_no not in range(1, 6):
        print("問題番号は1～5です")
        return

    # 問題ファイル読み込み
    wb = load_workbook(QUESTION_FILE)

    # 4. 受験者回答
    answer_sheet = wb["回答用紙"]
    user_answers = load_answers(answer_sheet)

    # 5. 正解シート
    answer_key_name = (
        f"{selected_type}_{selected_test_no}"
    )

    if answer_key_name not in wb.sheetnames:
        print(f"シート[{answer_key_name}]が存在しません")
        return

    answer_key_sheet = wb[answer_key_name]
    correct_answers = load_answers(answer_key_sheet)

    # 6. 採点
    total_questions = len(correct_answers)

    correct_count = 0

    for question_no, correct_answer in correct_answers.items():

        user_answer = user_answers.get(question_no)

        if user_answer == correct_answer:
            correct_count += 1

    wrong_count = total_questions - correct_count

    accuracy = 0

    if total_questions > 0:
        accuracy = (
            correct_count / total_questions
        ) * 100

    print("\n=== 採点結果 ===")
    print(f"正答数 : {correct_count}")
    print(f"誤答数 : {wrong_count}")
    print(f"正答率 : {accuracy:.2f}%")

    # 7. 結果シート出力
    result_wb = load_workbook(RESULT_FILE)
    result_ws = result_wb.active

    next_row = 2

    while result_ws.cell(next_row, 1).value:
        next_row += 1

    result_ws.cell(next_row, 1).value = datetime.now().strftime(
        "%Y/%m/%d"
    )
    result_ws.cell(next_row, 2).value = (
        selected_type if selected_type == "Listening" else ""
    )
    result_ws.cell(next_row, 3).value = (
        selected_type if selected_type == "Reading" else ""
    )
    result_ws.cell(next_row, 4).value = correct_count
    result_ws.cell(next_row, 5).value = round(accuracy, 2)

    result_wb.save(RESULT_FILE)

    print("\n結果シートへ出力しました。")


if __name__ == "__main__":
    main()