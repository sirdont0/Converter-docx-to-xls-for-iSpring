import os
import json
import zipfile
import shutil
import uuid
from copy import deepcopy
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox


# =========================
# НАСТРОЙКИ
# =========================

TEMPLATE_FOLDER = "./ready_example"
BUILD_DIR = "build_temp"


# =========================
# БЕЗОПАСНАЯ ЗАГРУЗКА JSON
# =========================

def safe_json_load(path):
    with open(path, "rb") as f:
        raw = f.read()

    for enc in ["utf-8", "utf-8-sig", "cp1251"]:
        try:
            return json.loads(raw.decode(enc))
        except:
            continue

    raise Exception("Не удалось определить кодировку: " + path)


def safe_json_save(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =========================
# WORD → ВОПРОСЫ
# =========================

def parse_table_from_docx(file_path):
    doc = Document(file_path)
    questions = []

    for table in doc.tables:
        for row in table.rows[1:]:
            cells = [cell.text.strip() for cell in row.cells]

            if len(cells) < 5:
                continue

            formulation = cells[2]
            variants = cells[3]
            correct = cells[4]

            if not formulation:
                continue

            questions.append({
                "text": formulation,
                "variants": variants,
                "correct": correct
            })

    return questions


# =========================
# ПОИСК КОНТЕЙНЕРА S
# =========================

def find_question_container(obj):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "S" and isinstance(v, list):
                if len(v) > 0 and isinstance(v[0], dict) and "tp" in v[0]:
                    return obj
            res = find_question_container(v)
            if res:
                return res
    elif isinstance(obj, list):
        for item in obj:
            res = find_question_container(item)
            if res:
                return res
    return None


# =========================
# СОЗДАНИЕ .QUIZ
# =========================

def create_quiz_package(word_file, parsed_questions):

    base_name = os.path.splitext(os.path.basename(word_file))[0]

    # очищаем временную папку
    if os.path.exists(BUILD_DIR):
        shutil.rmtree(BUILD_DIR)

    shutil.copytree(TEMPLATE_FOLDER, BUILD_DIR)

    document_path = os.path.join(BUILD_DIR, "document.json")
    metainfo_path = os.path.join(BUILD_DIR, "metainfo.json")

    document_data = safe_json_load(document_path)
    meta_data = safe_json_load(metainfo_path)

    container = find_question_container(document_data)
    if not container:
        raise Exception("Не найден контейнер S")

    template_questions = container["S"]

    mc_template = None
    ti_template = None

    for q in template_questions:
        if q.get("tp") == "MultipleChoice":
            mc_template = q
        if q.get("tp") == "TypeIn":
            ti_template = q

    if not mc_template:
        raise Exception("В шаблоне должен быть MultipleChoice вопрос")

    new_questions = []

    for item in parsed_questions:

        # =====================
        # MC / MR
        # =====================
        if item["variants"]:

            q = deepcopy(mc_template)
            q["i"] = uuid.uuid4().hex

            # чистим текст вопроса полностью
            q["D"]["d"] = [{
                "tp": "paragraph",
                "c": [{
                    "t": item["text"],
                    "tp": "text"
                }]
            }]

            lines = [v.strip() for v in item["variants"].split("\n") if ")" in v]
            correct_clean = item["correct"].replace("; ", ", ")
            correct_letters = [c.strip() for c in correct_clean.split(", ") if c.strip()]

            new_choices = []

            for line in lines:
                letter = line[0]
                text = line.split(")", 1)[1].strip()

                new_choices.append({
                    "i": uuid.uuid4().hex,
                    "t": {
                        "d": [{
                            "tp": "paragraph",
                            "c": [{
                                "t": text,
                                "tp": "text"
                            }]
                        }]
                    },
                    "c": letter in correct_letters
                })

            q["C"]["chs"] = new_choices

            # тип вопроса
            if len(correct_letters) > 1:
                q["tp"] = "MultipleResponse"
            else:
                q["tp"] = "MultipleChoice"

            # 1 балл за вопрос
            q["s"]["e"]["pt"] = 1

            new_questions.append(q)

        # =====================
        # TI
        # =====================
        else:

            if not ti_template:
                continue

            q = deepcopy(ti_template)
            q["i"] = uuid.uuid4().hex

            q["D"]["d"] = [{
                "tp": "paragraph",
                "c": [{
                    "t": item["text"],
                    "tp": "text"
                }]
            }]

            answers = [
                a.strip() for a in
                item["correct"].replace(",", "\n").replace(";", "\n").split("\n")
                if a.strip()
            ]

            q["C"]["chs"] = []

            for ans in answers:
                q["C"]["chs"].append({
                    "i": uuid.uuid4().hex,
                    "t": ans
                })

            q["s"]["e"]["pt"] = 1

            new_questions.append(q)

    # полностью заменяем вопросы
    container["S"] = new_questions

    # =========================
    # МЕТАДАННЫЕ
    # =========================

    meta_data["title"] = base_name
    meta_data["quiz"]["gradeInfo"]["maxScore"] = str(len(new_questions))

    # =========================
    # СОХРАНЕНИЕ
    # =========================

    safe_json_save(document_path, document_data)
    safe_json_save(metainfo_path, meta_data)

    # =========================
    # УПАКОВКА
    # =========================

    output_path = os.path.join(
        os.path.dirname(word_file),
        base_name + ".quiz"
    )

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(BUILD_DIR):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, BUILD_DIR)
                z.write(full_path, rel_path)

    shutil.rmtree(BUILD_DIR)

    print("Создан:", output_path)


# =========================
# GUI
# =========================

def select_files():
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        filetypes=[("Word files", "*.docx")]
    )

    if not files:
        return

    for file in files:
        questions = parse_table_from_docx(file)
        create_quiz_package(file, questions)

    messagebox.showinfo("Готово", "Тесты успешно созданы!")


if __name__ == "__main__":
    select_files()