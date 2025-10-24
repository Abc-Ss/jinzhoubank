import re
import xlrd
import xlwt
import chardet
import warnings


# -------------------- 1. 本行报盘 --------------------
def LocalOffer(txt_path, excel_path):
    # 忽略xlwt的未来警告
    warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

    # 读取文件，尝试多种编码
    encodings = ['gbk', 'utf-8', 'cp936']
    lines = []
    for encoding in encodings:
        try:
            with open(txt_path, 'r', encoding=encoding) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue
    if not lines:
        raise ValueError("无法读取文件，请检查编码或路径")

    # 排除最后一行数据（汇总行）
    if len(lines) > 0:
        lines = lines[:-1]

    data = []
    unprocessed_lines = []

    # 遍历每一行数据
    for line_num, line in enumerate(lines, 1):
        line = line.strip()
        if not line:
            continue

        # 正则表达式匹配多空格分隔的字段
        pattern = re.compile(
            r'^(\S+)\s+'  # 字段1：固定前缀
            r'(\d+)\s+'  # 字段2：序号
            r'(\S+)\s+'  # 字段3：固定码
            r'(\S+)\s+'  # 字段4：卡号
            r'(.+?)\s+'  # 字段5：姓名/公司名（支持含空格）
            r'(\S+)\s+'  # 字段6：固定值1
            r'(\S+)\s+'  # 字段7：应处理金额
            r'(\S+)\s+'  # 字段8：中间码1
            r'(\S*)\s*'  # 字段9：中间码2（允许后面空格数量任意）
            r'(\S*)$'  # 字段10：备注（允许空）
        )

        match = pattern.match(line)
        if match:
            try:
                card_num = match.group(4)
                company = match.group(5).strip()
                amount_str = match.group(7)
                remark = match.group(9).strip()

                # 处理金额：有值则保留两位小数并转为字符串，空值则记为"0.00"
                if amount_str:
                    # 先转换为浮点数处理，再格式化为两位小数的字符串
                    amount_text = f"{round(float(amount_str) / 100, 2):.2f}"
                else:
                    amount_text = "0.00"  # 明确为字符串格式的两位小数

                data.append([company, card_num, amount_text, remark, None, None])
            except (ValueError, IndexError) as e:
                unprocessed_lines.append(f"行{line_num}：解析错误 - {str(e)}")
        else:
            unprocessed_lines.append(f"行{line_num}：未匹配格式")

    # 直接使用xlwt创建xls文件
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('数据')  # 直接创建工作表

    # 定义格式
    header_font = xlwt.Font()
    header_font.bold = True
    header_style = xlwt.XFStyle()
    header_style.font = header_font
    header_style.num_format_str = '@'  # 表头统一使用文本格式

    # 文本格式（所有列都使用文本格式）
    text_style = xlwt.XFStyle()
    text_style.num_format_str = '@'  # 强制文本格式

    # 表头数据
    headers = ['姓名\n(不超过60个字节)', '卡号', '应处理金额(必须小于1亿)',
               '备注(不超过12个字节)', '实处理金额', '处理标志']

    # 写入表头
    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header, header_style)

    # 写入数据
    for row_idx, row_data in enumerate(data, 1):  # 从1开始（跳过表头）
        for col_idx, value in enumerate(row_data):
            # 所有列都应用文本格式
            cell_value = str(value) if value is not None else ''
            worksheet.write(row_idx, col_idx, cell_value, text_style)

    # 保存文件
    workbook.save(excel_path)

# -------------------- 2. 本行回盘 --------------------
def LocalReply(txt_report_path, excel_reply_path, txt_reply_path):
    mess = ""
    # ---------- ①. 读 xls 建字典 ----------
    wb = xlrd.open_workbook(excel_reply_path)
    sheet = wb.sheet_by_index(0)
    xls_map = {}
    for r in range(1, sheet.nrows):
        row = sheet.row_values(r)
        key = (str(row[0]).strip(),  # 姓名
               str(row[1]).strip(),  # 卡号
               str(int(round(float(row[2]) * 100))),  # 金额*100
               str(row[3]).strip())  # 原备注
        flag = str(row[5]).strip()
        xls_map[key] = flag          # value 先保留，后面还要用

    # ---------- ②. 编码侦探 ----------
    with open(txt_report_path, 'rb') as f:
        raw = f.read(10000)
    enc = chardet.detect(raw)['encoding'] or 'gbk'

    # ---------- ③. 收集 txt 中的 key ----------
    txt_keys = set()
    with open(txt_report_path, 'rb') as fin:
        for line_b in fin:
            line_u = line_b.decode(enc).rstrip('\r\n')
            parts_u = line_u.split()
            if len(parts_u) != 10:
                continue
            key = (parts_u[4], parts_u[3], parts_u[6], parts_u[8])
            txt_keys.add(key)

    # ---------- ④. 核对并逐行处理----------
    # 正则：10 列，第 9 列正好是 4 位数字
    # 分组：前面部分、第 9 列、后面部分
    xls_keys = set(xls_map.keys())
    if txt_keys == xls_keys:          # 集合相等：元素个数与内容完全一致
        mess = "文件信息一致"
        pat = re.compile(rb'^((?:\S+\s+){8})(\d{4})(\s+\S+.*)$')
        with open(txt_report_path, 'rb') as fin, open(txt_reply_path, 'wb') as fout:
            for line_b in fin:
                line_u = line_b.decode(enc).rstrip('\r\n')
                parts_u = line_u.split()
                if len(parts_u) != 10:  # 格式不对，原样输出
                    fout.write(line_b)
                    continue

                # 构造查询 key
                key = (parts_u[4],  # 姓名
                       parts_u[3],  # 卡号
                       parts_u[6],  # 金额
                       parts_u[8])  # 原 4 位备注

                if key in xls_map:
                    flag = xls_map[key]
                    new_note = ('001' if flag == '全部成功' else '002') + key[3]

                    # 用正则只替换第 9 列那 4 位数字，空格原样保留
                    def repl(m):
                        return m.group(1)[:-3] + new_note.encode(enc) + m.group(3)

                    line_b = pat.sub(repl, line_b)

                fout.write(line_b)
        print('回盘文件转换成功！', txt_reply_path)
        return mess,[],[]
    else:
        mess = "报盘txt与回盘xls信息不一致"
        # 如需详细差异，可打印：
        txt_xls = sorted(txt_keys - xls_keys, key=lambda x: (x[0], x[1]))
        xls_txt = sorted(xls_keys - txt_keys, key=lambda x: (x[0], x[1]))
        return mess,txt_xls,xls_txt         # 不一致可直接退出，不再生成回盘文件

# -------------------- 3. 他行报盘 --------------------

def OtherOffer(txt_path, excel_path):
    XLS_FIELDS = [
        "姓名\n(不超过60个字节)", "卡号", "行别", "跨行行号", "业务种类",
        "协议书号", "账号地址", "应处理金额(必须小于1亿)",
        "备注(不超过12个字节)", "实处理金额", "处理标志"
    ]
    data = []
    try:
        with open(txt_path, "r", encoding="gbk", errors="ignore") as f:
            for line_num, line in enumerate(f.readlines(), 1):
                line = line.rstrip("\n")
                if "天津泰达津联自来水有限公司" in line or line.strip() == "":
                    continue

                valid_blocks = list(filter(lambda x: x.strip() != "", line.split()))
                if len(valid_blocks) != 10:
                    continue

                # 数据映射逻辑（重点修改金额列处理）
                biz_type = valid_blocks[0][-5:].strip() if len(valid_blocks[0]) >= 5 else "00201"
                real_bank_code = valid_blocks[2][-12:].strip() if len(valid_blocks[2]) >= 12 else ""
                card_no = valid_blocks[3].strip()
                company_name = valid_blocks[4].strip()
                bank_type = valid_blocks[5].strip() if valid_blocks[5].strip() else "1"

                # 金额处理：先计算浮点数，再格式化为带两位小数的字符串（核心修改）
                raw_amount = valid_blocks[6].strip()
                if raw_amount.isdigit():
                    real_amount = round(int(raw_amount) / 100, 2)  # 计算金额（浮点数）
                    amount_str = f"{real_amount:.2f}"  # 格式化为两位小数的字符串（如"31.60"）
                else:
                    amount_str = "0.00"  # 异常情况默认值

                agreement_no = valid_blocks[7].strip()
                remark = valid_blocks[8].strip()

                # 组装行数据（金额列使用格式化后的字符串）
                xls_row = [
                    company_name, card_no, bank_type, real_bank_code, biz_type,
                    agreement_no, "", amount_str, remark, "", ""
                ]
                data.append(xls_row)

        if not data:
            return
    except Exception as e:
        return

    try:
        # 创建工作簿和工作表
        book = xlwt.Workbook(encoding='utf-8')
        sheet = book.add_sheet('sheet1')

        # 定义文本格式（所有列都用文本格式）
        text_style = xlwt.XFStyle()
        text_style.num_format_str = '@'  # Excel文本格式标识

        # 写入表头（文本格式）
        for col_idx, header in enumerate(XLS_FIELDS):
            sheet.write(0, col_idx, header, text_style)

        # 写入数据行（所有列都用文本格式，金额已提前格式化为两位小数字符串）
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, value in enumerate(row_data):
                # 所有值都按文本写入（金额已是带两位小数的字符串）
                sheet.write(row_idx, col_idx, str(value) if value != "" else "", text_style)

        # 保存文件
        book.save(excel_path)
    except Exception as e:
        return
    print('报盘文件转换成功！', excel_path)

# -------------------- 4. 他行回盘 --------------------
def OtherReply(txt_report_path, excel_reply_path,txt_reply_path):
    mess = ""
    # ---------- ①. 读 xls 建字典 ----------
    wb = xlrd.open_workbook(excel_reply_path)
    sheet = wb.sheet_by_index(0)
    xls_map = {}
    for r in range(1, sheet.nrows):
        row = sheet.row_values(r)
        key = (str(row[0]).strip(),  # 姓名
               str(row[1]).strip(),  # 卡号
               str(row[5]).strip(),  # 协议书号
               str(int(round(float(row[7]) * 100))),  # 金额*100
               str(row[8]).strip())  # 原备注
        flag = str(row[10]).strip()
        xls_map[key] = flag  # value 先保留，后面还要用

    # ---------- ②. 编码侦探 ----------
    with open(txt_report_path, 'rb') as f:
        raw = f.read(10000)
    enc = chardet.detect(raw)['encoding'] or 'gbk'

    # ---------- ③. 收集 txt 中的 key ----------
    txt_keys = set()
    with open(txt_report_path, 'rb') as fin:
        for line_b in fin:
            line_u = line_b.decode(enc).rstrip('\r\n')
            parts_u = line_u.split()
            if len(parts_u) != 10:
                continue
            key = (parts_u[4], parts_u[3], parts_u[7], parts_u[6], parts_u[8])
            txt_keys.add(key)

    # ---------- ④. 核对并逐行处理----------
    # 正则：10 列，第 9 列正好是 4 位数字
    # 分组：前面部分、第 9 列、后面部分
    xls_keys = set(xls_map.keys())
    if txt_keys == xls_keys:  # 集合相等：元素个数与内容完全一致
        mess = "文件信息一致"
        pat = re.compile(rb'^((?:\S+\s+){8})(\d{4})(\s+\S+.*)$')
        with open(txt_report_path, 'rb') as fin, open(txt_reply_path, 'wb') as fout:
            for line_b in fin:
                line_u = line_b.decode(enc).rstrip('\r\n')
                parts_u = line_u.split()
                if len(parts_u) != 10:  # 格式不对，原样输出
                    fout.write(line_b)
                    continue

                # 构造查询 key
                key = (parts_u[4],  # 姓名
                       parts_u[3],  # 卡号
                       parts_u[7],  # 协议书号
                       parts_u[6],  # 金额
                       parts_u[8])  # 原 4 位备注

                if key in xls_map:
                    flag = xls_map[key]
                    new_note = ('001' if flag == '全部成功' else '002') + key[4]

                    # 用正则只替换第 9 列那 4 位数字，空格原样保留
                    def repl(m):
                        return m.group(1)[:-3] + new_note.encode(enc) + m.group(3)

                    line_b = pat.sub(repl, line_b)

                fout.write(line_b)
        print('回盘文件转换成功！', txt_reply_path)
        return mess, [], []
    else:
        mess = "报盘txt与回盘xls信息不一致"
        # 如需详细差异，可打印：
        txt_xls = sorted(txt_keys - xls_keys, key=lambda x: (x[0], x[1]))
        xls_txt = sorted(xls_keys - txt_keys, key=lambda x: (x[0], x[1]))
        return mess, txt_xls, xls_txt  # 不一致可直接退出，不再生成回盘文件
