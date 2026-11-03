# -*- coding: utf-8 -*-
"""从需求规格书附录 C 抽取修订记录，生成 CHANGELOG.md（避免两处维护）"""
import io, re, sys
import docx

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import glob


def vkey(path):
    m = re.search(r'_v([\d.]+)\.docx$', path)
    return tuple(int(x) for x in m.group(1).split('.'))


# 自动取工作区里版本号最大的正本，避免每版手改路径
# 正本归档在 release/vX.Y/ 下，根目录也可能有工作副本，两处一起找
CANDIDATES = (glob.glob('核心系统迁移知识库需求规格书_v*.docx')
              + glob.glob('release/*/核心系统迁移知识库需求规格书_v*.docx'))
SRC = max(CANDIDATES, key=vkey)
OUT = 'CHANGELOG.md'
print('正本:', SRC)

d = docx.Document(SRC)
ps = d.paragraphs
start = [i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and '附录 C' in p.text][0]

blocks, cur = [], None
for p in ps[start + 1:]:
    t = p.text.strip()
    if not t:
        continue
    m = re.match(r'^(v[\d.]+)\s*修订[\u3000\s]*(.*)$', t)
    if m:
        cur = {'ver': m.group(1), 'head': m.group(2).rstrip('：:'), 'items': []}
        blocks.append(cur)
    elif cur is not None:
        if t.startswith('·'):
            cur['items'].append(t.lstrip('· ').strip())
        elif not cur['items']:
            cur['items'].append(t)

# 早期版本（v1.x）把整段写在标题行、无 · 条目：切出首句作标题，其余作正文
for b in blocks:
    if not b['items'] and ('：' in b['head'] or '。' in b['head']):
        parts = re.split(r'[：。]', b['head'], maxsplit=1)
        if len(parts) == 2 and parts[1].strip():
            b['head'], b['items'] = parts[0], [parts[1].strip()]

# 规模演进
scale = {}
for b in blocks:
    for it in b['items']:
        mm = re.search(r'需求总数由\s*(\d+)\s*条[^\d]{1,3}(\d+)\s*条', it)
        if mm:
            scale[b['ver']] = (int(mm.group(1)), int(mm.group(2)))
            break

# 未写明条数的中间版本：若更晚的版本给出了起点，回填为「不变」（限 v3.0 及之后）
# 最新版本没有「更晚的版本」可参照，改用上一版（idx+1）的终点值
vers = [b['ver'] for b in blocks]          # 倒序：新 → 旧
for idx, ver in enumerate(vers):
    if ver in scale:
        continue
    if float(ver[1:].split('.')[0]) < 3:
        continue
    later = [v for v in vers[:idx] if v in scale]   # 比它新的版本
    if later:
        base = scale[later[-1]][0]
    elif idx + 1 < len(vers) and vers[idx + 1] in scale:
        base = scale[vers[idx + 1]][1]
    else:
        continue
    scale[ver] = (base, base)

lines = []
lines.append('# 修订说明\n')
lines.append('按版本倒序记录改动要点。各条记录中提到的表号与条目编号，均指**该版本当时的状态**，'
             '跨版本引用需对照该版本的正文。\n')
lines.append('本文件由 `build/tools/gen_changelog.py` 从需求规格书附录 C 抽取生成，'
             '与正文同源，请勿手工编辑。\n')

rows = [(v, scale[v]) for v in vers if v in scale]
if rows:
    lines.append('## 规模演进\n')
    lines.append('| 版本 | 需求条数 | 较上一版 |')
    lines.append('|---|---|---|')
    for ver, (a, b) in rows:
        delta = b - a
        note = '不变' if delta == 0 else f'{a} → {b}（{"+" if delta > 0 else ""}{delta}）'
        lines.append(f'| {ver} | {b} | {note} |')
    lines.append('')

lines.append('---\n')

for b in blocks:
    lines.append(f'## {b["ver"]} — {b["head"]}\n')
    for it in b['items']:
        lines.append(f'- {it}')
    lines.append('')

io.open(OUT, 'w', encoding='utf-8').write('\n'.join(lines).rstrip() + '\n')
print('版本块', len(blocks), '| 规模行', len(rows))
