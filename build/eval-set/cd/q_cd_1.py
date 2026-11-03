# -*- coding: utf-8 -*-
"""01 开发与迁移（30 题）+ 02 业务分析（24 题）"""
from q_cd_base import (Q, pts, CAT_DEV, CAT_BIZ, T1, T2, T3, T4, T5, T6, T8)

G1 = []

# ---------------------------------------------------------------- 精准命中 12
G1.append(Q(CAT_DEV, "账户批量处理",
            "CBACT01C 这个程序的 PROGRAM-ID 是什么，定义在第几行？",
            T1, "简单", "是", "原始资料",
            pts("PROGRAM-ID 是 CBACT01C", "定义在 app/cbl/CBACT01C.cbl 第 23 行"),
            "app/cbl/CBACT01C.cbl:23", "代码直出", "可取用",
            "程序名 + 行号都要对上", "CBACT01C|23"))

G1.append(Q(CAT_DEV, "账务数据",
            "账户主记录一条多少字节？这个数是怎么算出来的？",
            T1, "中等", "是", "原始资料",
            pts("300 字节",
                "由 12 个字段加 FILLER(178) 合计而来，注释里也写明 RECLN 300",
                "与 acctdata.txt 每行 300 字符一致"),
            "app/cpy/CVACT01Y.cpy:2,5-17\napp/data/ASCII/acctdata.txt",
            "代码直出", "可取用",
            "300 + 至少一个构成依据", "300|178"))

G1.append(Q(CAT_DEV, "账务数据",
            "账户余额字段叫什么，能存多大、几位小数？",
            T1, "简单", "是", "原始资料",
            pts("ACCT-CURR-BAL",
                "PIC S9(10)V99，带符号，10 位整数加 2 位小数",
                "定长 13 字节（10+2 位数字 + 1 位符号）"),
            "app/cpy/CVACT01Y.cpy:7", "代码直出", "可取用",
            "字段名 + PIC 都要对", "ACCT-CURR-BAL|S9\\(10\\)V99|13"))

G1.append(Q(CAT_DEV, "在线交易入口",
            "在 CICS 里输交易码 CC00，执行的是哪个程序？这个绑定写在哪？",
            T1, "简单", "是", "原始资料",
            pts("COSGN00C，登录屏", "绑定写在 CARDDEMO.CSD 的 DEFINE TRANSACTION(CC00) 里，PROGRAM(COSGN00C)",
                "同时还有 DEFINE PROGRAM(COSGN00C) 的程序定义"),
            "app/csd/CARDDEMO.CSD:378-380\napp/csd/CARDDEMO.CSD:249",
            "代码直出", "可取用",
            "程序名 + 明确指出 CSD", "COSGN00C|CSD"))

G1.append(Q(CAT_DEV, "批量入账",
            "批量入账时判断信用卡额度够不够，这段在哪一行？超限记什么码？",
            T1, "中等", "是", "原始资料",
            pts("在 CBTRN02C 的 1500-B-LOOKUP-ACCT 段",
                "第 407 行 IF ACCT-CREDIT-LIMIT >= WS-TEMP-BAL 判断是否在额度内",
                "超限记 102，描述 'OVERLIMIT TRANSACTION'（410-412 行）"),
            "app/cbl/CBTRN02C.cbl:393-421", "代码直出", "可取用",
            "407 + 102 都要出现", "407|102"))

G1.append(Q(CAT_DEV, "交易数据",
            "交易金额字段 TRAN-AMT 的 PIC 是什么，占几个字节？",
            T1, "简单", "是", "原始资料",
            pts("PIC S9(09)V99", "带符号，9 位整数加 2 位小数，占 12 字节"),
            "app/cpy/CVTRA05Y.cpy:10", "代码直出", "可取用",
            "PIC + 12 都要对", "S9\\(09\\)V99|12"))

G1.append(Q(CAT_DEV, "利息计算",
            "利息是在哪个程序里算的，公式是什么？",
            T1, "中等", "是", "原始资料",
            pts("CBACT04C，批处理作业 INTCALC",
                "第 464-465 行：COMPUTE WS-MONTHLY-INT = (TRAN-CAT-BAL * DIS-INT-RATE) / 1200",
                "除以 1200 是把年利率百分数折成月息（年利率/100/12）"),
            "app/cbl/CBACT04C.cbl:462-470\napp/jcl/INTCALC.jcl",
            "代码直出", "可取用",
            "程序名 + 1200 都要出现", "CBACT04C|1200|DIS-INT-RATE"))

G1.append(Q(CAT_DEV, "客户数据",
            "客户主记录多少字节，客户姓名占多长？",
            T1, "简单", "是", "原始资料",
            pts("500 字节",
                "CUST-FIRST-NAME / CUST-MIDDLE-NAME / CUST-LAST-NAME 各 X(25)",
                "FILLER X(168) 补齐到 500"),
            "app/cpy/CVCUS01Y.cpy:2,6-8,23", "代码直出", "可取用",
            "500 + X(25) 都要对", "500|X\\(25\\)"))

G1.append(Q(CAT_DEV, "客户数据",
            "客户的信用评分存在哪个字段，几位？",
            T1, "简单", "是", "原始资料",
            pts("CUST-FICO-CREDIT-SCORE", "PIC 9(03)，3 位数字，无符号"),
            "app/cpy/CVCUS01Y.cpy:22", "代码直出", "可取用",
            "字段名 + 9(03)", "FICO|9\\(03\\)"))

G1.append(Q(CAT_DEV, "用户安全",
            "用户安全文件一条记录多长？密码和用户类型各占几位？",
            T1, "中等", "是", "原始资料",
            pts("SEC-USER-DATA 共 80 字节（8+20+20+8+1+23）",
                "SEC-USR-PWD PIC X(08)，8 位明文",
                "SEC-USR-TYPE PIC X(01)，取值 'A' 表示管理员"),
            "app/cpy/CSUSR01Y.cpy:18-24\napp/cpy/COCOM01Y.cpy:27",
            "代码直出", "可取用",
            "80 + X(08) + 'A'", "80|X\\(08\\)|A"))

G1.append(Q(CAT_DEV, "日期工具",
            "CSUTLDTC 这个程序是干什么的，谁调用它？",
            T1, "中等", "是", "原始资料",
            pts("日期合法性校验工具程序，PROGRAM-ID 在第 20 行",
                "COTRN02C 用它校验 Orig Date 和 Proc Date（第 393、413 行 CALL）",
                "返回 SEV-CD，'0000' 表示通过；MSG-NUM 为 2513 的情况被特殊放过"),
            "app/cbl/CSUTLDTC.cbl:20\napp/cbl/COTRN02C.cbl:389-427",
            "代码直出", "可取用",
            "日期校验 + CALL 方 + 2513", "日期|COTRN02C|2513"))

G1.append(Q(CAT_DEV, "数据导出",
            "CBEXPORT 的 PROGRAM-ID 定义在第几行？它和哪个程序配对使用？",
            T1, "简单", "是", "原始资料",
            pts("第 2 行", "与 CBIMPORT 配对，一导一出"),
            "app/cbl/CBEXPORT.cbl:2\napp/cbl/CBIMPORT.cbl",
            "代码直出", "可取用",
            "行号 2 + CBIMPORT", "2|CBIMPORT"))

# ---------------------------------------------------------------- 同义改写 3
G1.append(Q(CAT_DEV, "账务数据",
            "客户欠了多少钱存在哪个字段里？",
            T2, "简单", "是", "原始资料",
            pts("ACCT-CURR-BAL，账户当前余额",
                "PIC S9(10)V99，带符号，所以能表示欠款（正数=欠款，负数=溢缴）",
                "同一记录里还有 ACCT-CURR-CYC-CREDIT / DEBIT 记本周期发生额"),
            "app/cpy/CVACT01Y.cpy:7,13,14", "代码直出", "可取用",
            "ACCT-CURR-BAL + 带符号", "ACCT-CURR-BAL|S9"))

G1.append(Q(CAT_DEV, "交易数据",
            "日常交易文件和交易主文件的记录结构一样吗？",
            T2, "中等", "是", "原始资料",
            pts("字段完全对应，都是 350 字节",
                "只是前缀不同：DALYTRAN-* 与 TRAN-*",
                "入账时由 2000-POST-TRANSACTION 段逐字段搬运（425-438 行）"),
            "app/cpy/CVTRA05Y.cpy\napp/cpy/CVTRA06Y.cpy\napp/cbl/CBTRN02C.cbl:425-438",
            "代码直出", "可取用",
            "350 + 逐字段搬运", "350|DALYTRAN"))

G1.append(Q(CAT_DEV, "利息计算",
            "系统怎么知道该给某个账户按什么利率算息？",
            T2, "中等", "是", "原始资料",
            pts("利率来自披露组文件 DISGRP 的 DIS-INT-RATE",
                "键是 账户组 + 交易类型 + 类别（DIS-GROUP-KEY）",
                "读不到时用默认披露组（DEFAULT-DISCLOSURE-GROUP）"),
            "app/cpy/CVTRA02Y.cpy:5-9\napp/cbl/CBACT04C.cbl:462-467",
            "代码直出", "可取用",
            "DIS-INT-RATE + 键的三段", "DIS-INT-RATE|账户组"))

# ---------------------------------------------------------------- 多跳 9
G1.append(Q(CAT_DEV, "账务数据",
            "账户记录和客户记录加起来多长？",
            T3, "简单", "是", "原始资料",
            pts("300 + 500 = 800 字节", "两段分别来自 CVACT01Y 和 CVCUS01Y 的 01 层"),
            "app/cpy/CVACT01Y.cpy:2\napp/cpy/CVCUS01Y.cpy:2",
            "代码直出", "可取用",
            "800", "800"))

G1.append(Q(CAT_DEV, "批量入账",
            "一笔日常交易成功入账，会更新几个文件？分别是什么？",
            T3, "困难", "是", "原始资料",
            pts("三个：ACCOUNT-FILE 账户主、TCATBAL-FILE 交易类别余额、TRANSACT-FILE 交易主",
                "顺序是 2700-UPDATE-TCATBAL → 2800-UPDATE-ACCOUNT-REC → 2900-WRITE-TRANSACTION-FILE",
                "被拒时不走这三步，改写 DALYREJS"),
            "app/cbl/CBTRN02C.cbl:440-442", "代码直出", "可取用",
            "三个文件名 + 顺序", "TCATBAL|ACCOUNT|TRANSACT"))

G1.append(Q(CAT_DEV, "批量入账",
            "判断一笔交易是否超限，用到账户上哪几个字段？",
            T3, "困难", "是", "原始资料",
            pts("ACCT-CREDIT-LIMIT 信用额度",
                "ACCT-CURR-CYC-CREDIT 本周期贷记累计",
                "ACCT-CURR-CYC-DEBIT 本周期借记累计",
                "三者算出临时余额 WS-TEMP-BAL 后再与额度比较（403-407 行）"),
            "app/cbl/CBTRN02C.cbl:403-413", "代码直出", "可取用",
            "三个字段名都要出现", "CREDIT-LIMIT|CYC-CREDIT|CYC-DEBIT"))

G1.append(Q(CAT_DEV, "利息计算",
            "算一次利息要读哪些文件？",
            T3, "困难", "是", "原始资料",
            pts("TCATBALF 交易类别余额，提供计息基数",
                "DISCGRP 披露组，提供利率",
                "XREFFILE 卡号到账户的对照",
                "ACCTFILE 账户主；TRANFILE 用于写出利息交易"),
            "app/cbl/CBACT04C.cbl:62-95,462-500",
            "代码直出", "可取用",
            "TCATBALF + DISCGRP + XREF", "TCATBALF|DISCGRP|XREF"))

G1.append(Q(CAT_DEV, "在线交易入口",
            "从登录成功到进入「添加交易」屏，中间经过哪几个程序？",
            T3, "困难", "是", "原始资料",
            pts("COSGN00C 校验通过后 XCTL 到 COMEN01C（普通用户）或 COADM01C（管理员）",
                "主菜单选 CT02 再 XCTL 到 COTRN02C",
                "两次跳转都是 EXEC CICS XCTL，带 COMMAREA 传上下文"),
            "app/cbl/COSGN00C.cbl:230-239\napp/csd/CARDDEMO.CSD:439-441",
            "代码直出", "可取用",
            "COSGN00C + COMEN01C + COTRN02C", "COSGN00C|COMEN01C|COTRN02C"))

G1.append(Q(CAT_DEV, "账单支付",
            "在线做一笔账单支付，写出的交易记录里类型码和类别码是什么？",
            T3, "中等", "是", "原始资料",
            pts("TRAN-TYPE-CD = '02'（Payment）",
                "TRAN-CAT-CD 送的是 2，字段是 9(04)，实际存成 0002",
                "描述 'BILL PAYMENT - ONLINE'，来源 'POS TERM'，商户名 'BILL PAYMENT'"),
            "app/cbl/COBIL00C.cbl:219-229", "代码直出", "可取用",
            "02 + 0002 + BILL PAYMENT", "02|0002|BILL PAYMENT"))

G1.append(Q(CAT_DEV, "批量入账",
            "POSTTRAN 这个作业里，CBTRN02C 用了哪几个 DD，分别指向什么数据集？",
            T3, "困难", "是", "原始资料",
            pts("TRANFILE → AWS.M2.CARDDEMO.TRANSACT.VSAM.KSDS",
                "DALYTRAN → AWS.M2.CARDDEMO.DALYTRAN.PS",
                "XREFFILE → AWS.M2.CARDDEMO.CARDXREF.VSAM.KSDS",
                "ACCTFILE、TCATBALF 分别指向账户主与类别余额；DALYREJS 是新建 GDG(+1)"),
            "app/jcl/POSTTRAN.jcl", "代码直出", "可取用",
            "三个 DD 名 + 至少两个数据集名", "TRANFILE|DALYTRAN|XREFFILE"))

G1.append(Q(CAT_DEV, "客户数据",
            "如果把客户地址第三行从 X(50) 扩到 X(60)，客户记录会变成多少字节？还要动哪里？",
            T3, "困难", "是", "原始资料",
            pts("记录会从 500 变成 510 字节（FILLER 不动的话）",
                "要保住 500，得把 FILLER 从 X(168) 减到 X(158)",
                "同时要改对应数据集的 LRECL，以及已有的 custdata 数据文件（50 条 × 500）"),
            "app/cpy/CVCUS01Y.cpy:11,23\napp/data/ASCII/custdata.txt",
            "代码直出", "可取用",
            "510 或 158 至少答对一个 + 提到 LRECL 或数据文件",
            "510|158"))

G1.append(Q(CAT_DEV, "批量入账",
            "批量入账和在线账单支付都改账户余额，改法一样吗？",
            T3, "困难", "是", "原始资料",
            pts("批量：ADD 金额 TO ACCT-CURR-BAL，再按正负累加 CURR-CYC-CREDIT/DEBIT，并更新 TCATBAL",
                "在线还款：COMPUTE ACCT-CURR-BAL = ACCT-CURR-BAL - TRAN-AMT，只改这一个字段",
                "在线这条不动周期累计，也不更新类别余额，两条路径口径不同"),
            "app/cbl/CBTRN02C.cbl:547-552\napp/cbl/COBIL00C.cbl:234-235",
            "代码直出", "有争议",
            "必须点出「只改余额」这个差异", "CURR-CYC|只改余额|TCATBAL",
            note="真冲突（同一账户同一属性，两条路径更新范围不同）"))

G1.append(Q(CAT_DEV, "交易数据",
            "日常交易文件 DALYTRAN 一条记录多少字节？和交易主文件一样吗？",
            T1, "简单", "是", "原始资料",
            pts("350 字节，和 TRAN-RECORD 完全相同",
                "字段从 DALYTRAN-ID 到 FILLER X(20) 逐个对应",
                "数据文件 dailytran.txt 共 300 行、每行 350 字符"),
            "app/cpy/CVTRA06Y.cpy\napp/data/ASCII/dailytran.txt",
            "代码直出", "可取用",
            "350 + 300 行", "350|300"))

G1.append(Q(CAT_DEV, "交易类别余额",
            "交易类别余额记录的主键是哪三段？顺序是什么？",
            T3, "中等", "是", "原始资料",
            pts("TRANCAT-ACCT-ID 9(11) 账户号",
                "TRANCAT-TYPE-CD X(02) 交易类型",
                "TRANCAT-CD 9(04) 交易类别",
                "13+2+4 = 19 字节键，后接 TRAN-CAT-BAL S9(09)V99 与 FILLER X(22)"),
            "app/cpy/CVTRA01Y.cpy:5-10", "代码直出", "可取用",
            "三段字段名 + 顺序", "ACCT-ID|TYPE-CD|CAT-CD"))

G1.append(Q(CAT_DEV, "批量入账",
            "一批交易里有被拒绝的，作业结束时返回码是多少？为什么这么设？",
            T6, "中等", "是", "原始资料",
            pts("返回码 4（第 229-231 行：IF WS-REJECT-COUNT > 0 MOVE 4 TO RETURN-CODE）",
                "这样调度系统能区分「跑完但有拒绝」和「完全成功」，不必翻日志",
                "全部成功时不设返回码，即 0"),
            "app/cbl/CBTRN02C.cbl:229-231", "代码直出", "可取用",
            "4 + 有拒绝", "4|REJECT-COUNT"))

# ---------------------------------------------------------------- 跨语言 3
G1.append(Q(CAT_DEV, "ZOpenEditor 样本",
            "ZOpenEditor 里 CUSTCOPY 定义的客户记录多长？怎么验证这个数？",
            T2, "中等", "是", "原始资料",
            pts("387 字节",
                "与 RESOURCES/SAMPLE.CUSTFILE.txt 的实际行长完全一致，可逐行核",
                "同一份 copybook 里还有一条联系人记录 69 字节"),
            "COPYBOOK/CUSTCOPY.cpy:24-42\nRESOURCES/SAMPLE.CUSTFILE.txt",
            "代码直出", "可取用",
            "387 + 提到数据文件验证", "387"))

G1.append(Q(CAT_DEV, "ZOpenEditor 样本",
            "汇编程序 ASAM1 的 CSECT 叫什么，定义在第几行？输入输出 DD 名是什么？",
            T2, "中等", "是", "原始资料",
            pts("CSECT 名 ASAM1，第 2 行定义",
                "输入 DD 名 FILEIN，输出 FILEOUT",
                "两者都是 QSAM，程序内 OPEN 在 98-101 行"),
            "ASM/ASAM1.asm:2,25-27,98-101", "代码直出", "可取用",
            "ASAM1 + FILEIN + FILEOUT", "ASAM1|FILEIN|FILEOUT"))

G1.append(Q(CAT_DEV, "ZOpenEditor 样本",
            "SAMPLE.TRANFILE.txt 里各行的长度是多少？为什么不一样？",
            T2, "中等", "是", "原始资料",
            pts("依次是 51、13、13、52 字符",
                "不同交易类型用不同记录长度，不是定长文件",
                "而 CUSTFILE 是定长 387/69 两种，同样是变长组织方式"),
            "RESOURCES/SAMPLE.TRANFILE.txt", "代码直出", "可取用",
            "51|13|52", "51|13|52"))

# ================================================================ 02 业务分析
G2 = []

G2.append(Q(CAT_BIZ, "额度控制",
            "判断一笔刷卡会不会超限，系统用的是账户当前余额，还是本周期累计？",
            T3, "困难", "是", "原始资料",
            pts("用本周期累计：WS-TEMP-BAL = ACCT-CURR-CYC-CREDIT - ACCT-CURR-CYC-DEBIT + 本笔金额",
                "不是直接用 ACCT-CURR-BAL",
                "所以周期累计没清零之前，额度判断会一直带着历史累计，两者在周期切换点会分叉"),
            "app/cbl/CBTRN02C.cbl:403-407", "代码直出", "可取用",
            "必须答「本周期累计」并点出三个字段", "本周期|CYC"))

G2.append(Q(CAT_BIZ, "利息计算",
            "利息是按哪个金额算的？按账户总余额还是分类余额？",
            T3, "困难", "是", "原始资料",
            pts("按 TRAN-CAT-BAL，即交易类别余额，一个账户的每个类型+类别组合一条",
                "利息 = (TRAN-CAT-BAL × DIS-INT-RATE) / 1200",
                "不是按账户总余额一次性算"),
            "app/cbl/CBACT04C.cbl:462-467", "代码直出", "可取用",
            "TRAN-CAT-BAL + 分类", "TRAN-CAT-BAL|类别"))

G2.append(Q(CAT_BIZ, "利息计算",
            "利率能不能按单个客户单独设？",
            T3, "中等", "是", "原始资料",
            pts("不能按客户设。利率键是 账户组 + 交易类型 + 类别",
                "账户组字段是 ACCT-GROUP-ID，不是客户号",
                "同一组内同类型同类别的账户共用一个利率"),
            "app/cpy/CVTRA02Y.cpy:5-9\napp/cpy/CVACT01Y.cpy:16",
            "代码直出", "可取用",
            "账户组 + 不是按客户", "账户组|ACCT-GROUP-ID"))

G2.append(Q(CAT_BIZ, "费用",
            "系统收不收手续费？",
            T6, "困难", "是", "原始资料",
            pts("代码里有 1400-COMPUTE-FEES 段，但段内只有注释 * To be implemented 和 EXIT",
                "等于没有实现，不产生任何手续费交易",
                "对外只能答「未实现」，不能答「收」或「不收」"),
            "app/cbl/CBACT04C.cbl:518-520", "代码直出", "可取用",
            "必须出现「未实现」或 To be implemented", "To be implemented|未实现",
            note="段落名不等于实现，答「收手续费」即为幻觉"))

G2.append(Q(CAT_BIZ, "账单支付",
            "客户能不能只还一部分账单？",
            T3, "中等", "是", "原始资料",
            pts("这个功能不能。还款金额直接取 ACCT-CURR-BAL 全额（第 224 行）",
                "账单支付屏上没有金额输入框，只有一个账户号和确认项",
                "代码里也没有部分还款的入参路径"),
            "app/cbl/COBIL00C.cbl:224\napp/bms/COBIL00.bms",
            "代码直出", "可取用",
            "全额 + 不能", "全额|不能|CURR-BAL"))

G2.append(Q(CAT_BIZ, "账单支付",
            "账户余额已经是 0 或负数，还能做还款吗？",
            T6, "中等", "是", "原始资料",
            pts("不能。条件是 ACCT-CURR-BAL <= ZEROS 且输入了账户号",
                "提示 'You have nothing to pay...'",
                "这是业务规则，不是故障"),
            "app/cbl/COBIL00C.cbl:197-205", "代码直出", "可取用",
            "nothing to pay + 小于等于零", "nothing to pay|ZEROS"))

G2.append(Q(CAT_BIZ, "冲正",
            "系统支持冲正吗？",
            T5, "困难", "是", "原始资料",
            pts("交易类型码表里有 06 Reversal",
                "但全库没有任何程序按 06 分支处理，也没有冲正/反向记账逻辑",
                "属于「码表有定义、功能未实现」，不是「不支持冲正」也不是「支持冲正」"),
            "app/data/ASCII/trantype.txt\napp/cbl/（全库检索无 06 处理）",
            "代码直出", "有争议",
            "06 + Reversal + 未实现", "06|Reversal|未实现",
            note="码表有、实现无，最容易被答成「支持」"))

G2.append(Q(CAT_BIZ, "交易分类",
            "系统里交易类型有几种，分别是什么？",
            T1, "简单", "是", "原始资料",
            pts("7 种：01 Purchase、02 Payment、03 Credit、04 Authorization、05 Refund、06 Reversal、07 Adjustment",
                "码表文件 7 行，每行 60 字符（2 位代码 + 50 位描述 + 8 位填充）"),
            "app/data/ASCII/trantype.txt\napp/cpy/CVTRA03Y.cpy",
            "代码直出", "可取用",
            "7 种 + 至少三个名称", "7|Purchase|Reversal"))

G2.append(Q(CAT_BIZ, "交易分类",
            "交易类别码有多少种？举三个例子。",
            T1, "简单", "是", "原始资料",
            pts("18 种，码表文件 18 行",
                "例如 010001 Regular Sales Draft、010005 Interest Amount、020001 Cash payment",
                "类别码是 4 位，与 2 位类型码组成 6 位键"),
            "app/data/ASCII/trancatg.txt\napp/cpy/CVTRA04Y.cpy",
            "代码直出", "可取用",
            "18 + 一个正确例子", "18|Regular Sales Draft|Interest Amount"))

G2.append(Q(CAT_BIZ, "账户状态",
            "账户有几种状态？能不能冻结？",
            T3, "中等", "是", "原始资料",
            pts("ACCT-ACTIVE-STATUS PIC X(01)，只有 Y/N 两种取值",
                "没有冻结、销户、止付等中间状态",
                "所以它不是状态机，只是一个开关"),
            "app/cpy/CVACT01Y.cpy:6\napp/cbl/COACTUPC.cbl:504",
            "代码直出", "可取用",
            "Y/N 两种 + 没有冻结", "Y|N|两种"))

G2.append(Q(CAT_BIZ, "卡片状态",
            "卡能不能单独停用？",
            T2, "简单", "是", "原始资料",
            pts("能。CARD-ACTIVE-STATUS PIC X(01) 在卡记录里",
                "它与账户状态 ACCT-ACTIVE-STATUS 是两个独立字段",
                "所以同一账户下可以停一张卡、留另一张"),
            "app/cpy/CVACT02Y.cpy:10\napp/cpy/CVACT01Y.cpy:6",
            "代码直出", "可取用",
            "CARD-ACTIVE-STATUS + 独立", "CARD-ACTIVE-STATUS"))

G2.append(Q(CAT_BIZ, "有效期控制",
            "卡过期了还能不能入账？",
            T3, "中等", "是", "原始资料",
            pts("不能。第 414 行比较 ACCT-EXPIRAION-DATE 与交易日期前 10 位",
                "过期记 103，描述 'TRANSACTION RECEIVED AFTER ACCT EXPIRATION'",
                "注意比较的是账户有效期，不是卡自身的 CARD-EXPIRAION-DATE"),
            "app/cbl/CBTRN02C.cbl:414-419", "代码直出", "可取用",
            "103 + 账户有效期（不是卡有效期）", "103|EXPIRAION"))

G2.append(Q(CAT_BIZ, "权限",
            "管理员和普通用户有什么不一样？",
            T3, "中等", "是", "原始资料",
            pts("登录成功按 SEC-USR-TYPE 分流：等于 'A' 走 COADM01C 管理菜单，否则走 COMEN01C",
                "管理菜单里多出用户管理和交易类型管理两块",
                "进入功能后各程序不再校验角色，权限只体现在入口这一跳"),
            "app/cbl/COSGN00C.cbl:230-239\napp/cpy/COCOM01Y.cpy:27",
            "代码直出", "可取用",
            "COADM01C + 'A' + 入口分流", "COADM01C|A"))

G2.append(Q(CAT_BIZ, "权限",
            "普通用户能不能直接敲交易码进管理菜单？",
            T3, "困难", "是", "原始资料",
            pts("能。CA00 在 CSD 里是 STATUS(ENABLED)，没有按用户类型附加限制",
                "登录时只是默认不跳转过去，不等于交易层拦截",
                "所以权限控制只发生在登录那一跳，交易本身不校验"),
            "app/csd/CARDDEMO.CSD:327-329\napp/cbl/COSGN00C.cbl:230-239",
            "代码直出", "可取用",
            "能 + 指出 CSD 未限制", "CA00|ENABLED|能"))

G2.append(Q(CAT_BIZ, "安全",
            "用户密码是怎么存的？",
            T1, "中等", "是", "原始资料",
            pts("SEC-USR-PWD PIC X(08)，明文",
                "登录时直接比较 IF SEC-USR-PWD = WS-USER-PWD（第 223 行）",
                "没有哈希、没有加盐、没有失败锁定"),
            "app/cpy/CSUSR01Y.cpy:21\napp/cbl/COSGN00C.cbl:223",
            "代码直出", "可取用",
            "明文 + X(08)", "明文|X\\(08\\)"))

G2.append(Q(CAT_BIZ, "数据关系",
            "一个客户能有几张卡？",
            T3, "困难", "是", "原始资料",
            pts("卡记录用 CARD-ACCT-ID 指向账户，卡号到客户/账户的对应放在 XREF",
                "XREF 一条记录一个卡号，所以同一账户可以挂多张卡，同一客户也可以有多张",
                "代码里没有数量上限校验"),
            "app/cpy/CVACT02Y.cpy:6\napp/cpy/CVACT03Y.cpy:5-8",
            "代码直出", "可取用",
            "通过 XREF + 无上限", "XREF|卡号"))

G2.append(Q(CAT_BIZ, "额度定义",
            "信用额度和取现额度是两个字段吗，有什么关系？",
            T1, "中等", "是", "原始资料",
            pts("两个独立字段：ACCT-CREDIT-LIMIT、ACCT-CASH-CREDIT-LIMIT，都是 S9(10)V99",
                "批量入账校验只用前者（407 行）",
                "取现额度要到授权模块才用到"),
            "app/cpy/CVACT01Y.cpy:8-9\napp/cbl/CBTRN02C.cbl:407",
            "代码直出", "可取用",
            "两个字段 + 入账只用前者", "CASH-CREDIT-LIMIT|两个"))

G2.append(Q(CAT_BIZ, "周期口径",
            "本周期消费累计什么时候清零？",
            T3, "困难", "是", "原始资料",
            pts("全库没有任何清零逻辑",
                "只有入账时累加（548-552 行），没有周期切换作业、也没有重置步骤",
                "所以字段名叫本周期，实际是从开户以来的累计"),
            "app/cbl/CBTRN02C.cbl:548-552\napp/jcl/（无清零作业）",
            "代码直出", "可取用",
            "没有清零 + 累计", "没有清零|累计",
            note="名字与实际行为不符，属文档未提及的实现事实"))

G2.append(Q(CAT_BIZ, "对账单",
            "生成对账单的作业读哪几个文件？",
            T3, "中等", "是", "原始资料",
            pts("作业 CREASTMT 跑 CBSTM03A",
                "按 XREF → CUSTFILE → ACCTFILE → TRNXFILE 的顺序读",
                "四个都是输入，输出是对账单"),
            "app/cbl/CBSTM03A.CBL:317-331\nREADME.md（Batch Components 表）",
            "代码直出", "可取用",
            "四个文件名", "XREF|CUSTFILE|ACCTFILE|TRNXFILE"))

G2.append(Q(CAT_BIZ, "报表口径",
            "交易报表上的总笔数，含不含被拒绝的交易？",
            T3, "中等", "是", "原始资料",
            pts("含。ADD 1 TO WS-TRANSACTION-COUNT 在读入之后、校验之前（第 206 行）",
                "被拒的另外累加 WS-REJECT-COUNT（第 214 行）",
                "所以总笔数 = 成功笔数 + 拒绝笔数"),
            "app/cbl/CBTRN02C.cbl:206,214", "代码直出", "可取用",
            "含 + 两个计数", "含|206"))

G2.append(Q(CAT_BIZ, "拒绝处理",
            "被拒绝的交易记录写到哪？带什么信息？",
            T1, "中等", "是", "原始资料",
            pts("写到 DALYREJS，作业里定义为 DALYREJS(+1) 的新建 GDG",
                "由 2500-WRITE-REJECT-REC 段写（446-465 行）",
                "同时带上 WS-VALIDATION-TRAILER，即失败原因码与描述"),
            "app/cbl/CBTRN02C.cbl:446-465\napp/jcl/POSTTRAN.jcl",
            "代码直出", "可取用",
            "DALYREJS + GDG", "DALYREJS|GDG"))

G2.append(Q(CAT_BIZ, "交易标识",
            "交易 ID 是怎么生成的？批量和在线一样吗？",
            T3, "困难", "是", "原始资料",
            pts("利息交易：PARM 传进来的日期 + 4 位序号，STRING PARM-DATE WS-TRANID-SUFFIX（473-480 行）",
                "在线还款：STARTBR 定位 HIGH-VALUES 后 READPREV 取最后一条，ID 加 1（212-217 行）",
                "两种方式不同，并发时可能撞号"),
            "app/cbl/CBACT04C.cbl:473-480\napp/cbl/COBIL00C.cbl:212-219",
            "代码直出", "可取用",
            "两种方式都要说 + 冲突风险", "PARM|加 1|两种"))

G2.append(Q(CAT_BIZ, "数据保留",
            "交易数据保留多久？有没有归档或清理策略？",
            T4, "困难", "否", "—",
            pts("库内没有任何交易归档或清理逻辑",
                "唯一的清理是授权过期清除（CBPAUP0C），与交易无关",
                "应明确回答「库内无此信息」"),
            "—（应无出处）", "—", "可取用",
            "必须承认查不到", "没有|归档|查不到",
            note="全库检索 archive/purge/retention 无结果，仅授权模块有 purge"))

G2.append(Q(CAT_BIZ, "总账",
            "利息收入在总账里怎么记？有没有会计分录？",
            T4, "困难", "否", "—",
            pts("库内没有总账、科目、分录相关定义",
                "利息只写一条交易记录（类型 01、类别 05），不产生分录",
                "要回答「需财务口径，库内无此信息」"),
            "—（应无出处）", "—", "可取用",
            "必须承认查不到", "没有|查不到|分录",
            note="CardDemo 是信用卡业务示例，不含总账子系统"))
