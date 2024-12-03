import re
from bs4 import BeautifulSoup
import os
import pandas as pd

def read_html_file(file_path):
    """读取HTML文件内容"""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

def read_question_bank(xlsx_path):
    """读取题库XLSX文件并转换为DataFrame"""
    try:
        df = pd.read_excel(xlsx_path, engine='openpyxl')  # 读取XLSX文件
        print(f"文件列名：{df.columns.tolist()}")  # 打印列名用于调试
    except Exception as e:
        print(f"读取文件失败：{e}")
        return None

    # 转换为题库格式
    questions = []
    for index, row in df.iterrows():
        # 通过DataFrame行获取题目及选项
        question = row['题面']
        options = [row['选项1'], row['选项2'], row['选项3'], row['选项4']]
        answer = row['答案']

        questions.append({
            'question': question,
            'options': options,
            'answer': answer
        })

    return questions, df


def parse_html_questions(html):
    """解析HTML中的所有题目内容"""
    soup = BeautifulSoup(html, 'html.parser')
    questions = []

    # 找到所有题目div
    question_divs = soup.find_all('div', class_='field ui-field-contain', type=['3'])

    for div in question_divs:
        question_div = div.find('div', class_='topichtml')
        if not question_div:
            continue

        question_text = question_div.get_text(strip=True)

        # 获取选项及其值
        options = []
        values = []  # 存储选项对应的value值
        for radio_div in div.find_all('div', class_='ui-radio'):
            input_tag = radio_div.find('input')
            if input_tag:
                value = input_tag.get('value', '')
                values.append(value)

            label_div = radio_div.find('div', class_='label')
            if label_div:
                options.append(label_div.get_text(strip=True))

        questions.append({
            'question': question_text,
            'options': options,
            'values': values
        })

    return questions

def similarity_score(text1, text2):
    """计算两个文本的相似度"""
    # 移除标点符号和空格
    text1 = re.sub(r'[^\w\s]', '', text1)
    text2 = re.sub(r'[^\w\s]', '', text2)

    # 转换为集合计算重合度
    set1 = set(text1)
    set2 = set(text2)

    intersection = len(set1.intersection(set2))
    union = len(set1.union(set2))

    return intersection / union if union > 0 else 0

def find_matching_options(html_options, bank_options):
    """找到选项的对应关系"""
    mapping = {}
    for i, html_opt in enumerate(html_options):
        max_sim = -1
        best_match = -1
        for j, bank_opt in enumerate(bank_options):
            sim = similarity_score(html_opt, bank_opt)
            if sim > max_sim:
                max_sim = sim
                best_match = j
        if max_sim > 0.5:  # 设置一个阈值
            mapping[i] = best_match
    return mapping

def convert_answer(original_answer, options, option_mapping):
    """转换答案从题库格式到网页格式"""
    if not original_answer:
        return None
    # 假设题库中答案是ABCD格式
    answer_index = ord(original_answer) - ord('A')
    # 在映射中找到对应的网页选项文本
    for html_idx, bank_idx in option_mapping.items():
        if bank_idx == answer_index:
            return options[html_idx]
    return None

def find_matching_question(html_question, question_bank, threshold=0.7):
    """在题库中查找匹配的题目"""
    best_match = None
    best_score = 0
    best_option_mapping = None

    html_q_text = html_question['question']

    for bank_q in question_bank:
        # 计算题目相似度
        score = similarity_score(html_q_text, bank_q['question'])
        if score > best_score and score >= threshold:
            # 找到选项的对应关系
            option_mapping = find_matching_options(html_question['options'], bank_q['options'])
            if option_mapping:  # 只有在能够匹配选项时才考虑这个题目
                best_score = score
                best_match = bank_q
                best_option_mapping = option_mapping

    return best_match, best_score, best_option_mapping

def process_questions(html_path, xlsx_path):
    """处理HTML文件中的所有题目并在题库中查找匹配"""
    html_content = read_html_file(html_path)
    question_bank, df = read_question_bank(xlsx_path)

    if not question_bank:
        return []

    html_questions = parse_html_questions(html_content)

    results = []
    unmatched_questions = []
    for html_question in html_questions:
        matching_question, score, option_mapping = find_matching_question(html_question, question_bank)

        if matching_question and option_mapping:
            # 转换答案
            web_answer = convert_answer(matching_question['answer'],
                                     html_question['options'],
                                     option_mapping)

            results.append({
                'found': True,
                'html_question': html_question['question'],
                'matching_question': matching_question['question'],
                'original_answer': matching_question['answer'],
                'web_answer': web_answer,
                'similarity_score': score,
                'option_mapping': option_mapping
            })
        else:
            unmatched_questions.append({
                'question': html_question['question'],
                'options': html_question['options'],
                'answer': '',
            })

    # 将未匹配的题目写入XLSX文件
    if unmatched_questions:
        unmatched_df = pd.DataFrame(unmatched_questions)
        unmatched_df.to_excel(xlsx_path, index=False, engine='openpyxl', mode='a')

    return results

def main():
    import sys
    import io

    html_path = 'RMUC 2025规则测评.html'
    xlsx_path = '完整题库 最低99最高100.xlsx'
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    if not os.path.exists(html_path):
        print(f"错误：找不到HTML文件：{html_path}")
        return
    if not os.path.exists(xlsx_path):
        print(f"错误：找不到题库文件：{xlsx_path}")
        return

    results = process_questions(html_path, xlsx_path)

    count = 0
    for result in results:
        if result['found']:
            print(f"题目 {count + 1}:")
            print(f"正确答案：{result['web_answer']}")
        else:
            print(f"题目 {count + 1}：未找到匹配")
            print(f"网页题目：{result['html_question']}")
        count += 1

        if (count % 5) == 0:
            print("\n" + "=" * 50 + "\n")

if __name__ == "__main__":
    main()
