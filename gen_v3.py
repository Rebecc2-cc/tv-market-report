# -*- coding: utf-8 -*-
"""
阶段 E：生成 8品牌AVC定位布局报告_v3.html（v3.1）
唯一写入目标：8品牌AVC定位布局报告_v3.html
保护名单内的文件绝不写入（2026-08-28 覆盖事故教训）

v3.1 相对 v3 的变更（2026-09-04，用户确认）：
1. 品牌口径统一为 8 品牌，产物改名 8品牌AVC定位布局报告_v3.html（旧 6品牌/8品牌 文件均保留不动）
2. 交互增强：URL 状态记忆 / 导出 CSV / 图1图2 一键展开收起 / 拼音首字母搜索 / Esc 关闭弹窗
3. 内容修正：图4「能效全缺」警告为过时信息（实际 42/71 款有能效数据），改为动态真实覆盖提示；
   footer 全部数字改为动态计算（原 60吋3款/小米红米52款均已过时）
4. 视觉：图1 热力刻度图例、对比表文字居中、各图缺参统计徽章
"""
import json, os, sys, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "master_v3.json")
OUT = os.path.join(BASE, "8品牌AVC定位布局报告_v3.html")

# ============ 写入保护断言 ============
PROTECTED = {
    "参数差异对比报告.html", "tv-comparison.html",
    "master.json", "report_data.json", "master_v2.json", "master_v3.json",
    "build_report.py", "gen_html.py", "build_master.py",
    "gen_v2.py", "build_v2.py", "build_v3.py",
    "8品牌AVC定位布局报告.html", "6品牌AVC定位布局报告.html",
    "电视参数一览表.xlsx", "电视参数一览表-整理后.xlsx",
    "小米红米_待补参数清单.xlsx",
}
_target = os.path.basename(OUT)
assert _target not in PROTECTED, f"[保护] 拒绝写入受保护文件：{_target}"
assert _target == "8品牌AVC定位布局报告_v3.html", f"[保护] 产物文件名异常：{_target}"
assert os.path.abspath(os.path.dirname(OUT)) == os.path.abspath(BASE), "[保护] 产物路径越界"
assert os.path.abspath(SRC) != os.path.abspath(OUT), "[保护] 输入与输出不能是同一文件"
print(f"[保护] 写入目标校验通过：{OUT}")

BRANDS = ["海信", "TCL", "创维", "Vidda", "雷鸟", "酷开", "小米", "红米"]
COLOR = {
    "海信": "#00AAA6", "TCL": "#ED1C24", "创维": "#0E7CE8",
    "Vidda": "#29B6F6", "雷鸟": "#1E3A8A", "酷开": "#F4511E",
    "小米": "#FF9500", "红米": "#C2185B",
}
# 4 个品牌展示产品名（雷鸟/Vidda/小米/红米）；其余品牌展示 AVC 原始型号名
BRANDS_USE_PRODUCT = {"雷鸟", "Vidda", "小米", "红米"}
# 图1 尺寸列（合并 100 吋+ 到 98 吋列，按用户要求改名为「98吋+」），共 9 列
FIG1_COLS = [32, 40, 43, 50, 55, 65, 75, 85, 98]
FIG1_LAB = {32: "32吋", 40: "40吋", 43: "43吋", 50: "50吋", 55: "55吋*",
            65: "65吋", 75: "75吋", 85: "85吋", 98: "98吋+"}
# 尺寸筛选档（沿用 7 档，与 size_bucket 一致）
SIZE_ORDER = ["50吋及以下", "55吋", "65吋", "75吋", "85吋", "98-100吋", "100吋以上"]
BELT_ORDER = ["<1000", "1000-1999", "2000-2999", "3000-4999", "5000-6999", "7000-7999",
              "8000-8999", "9000-9999", "10000-12999", "13000-15999", "16000-19999", "20000+"]
# 图3：分区档对齐参考（高档按分区数细分：三千级／四千-七千级／八千级+）；刷新率用具体 Hz 值（保留 132Hz 单独列；含全部 8 档）
PART_ORDER = ["八千级+", "四千-七千级", "三千级", "2000级", "千级", "百千级", "几百", "百级", "几十区", "<100", "无分区"]
# 图4：能效 × 分辨率（对齐参考）
MEM_ORDER = ["≤2GB", "3GB", "4GB+", "待补"]
RES_ORDER = ["4K", "1080P", "WXGA", "待补"]
ENERGY_ORDER = ["一级", "二级", "三级", "四级", "待补"]
# 内存颜色：单个卡片底色（拉大色阶，浅→深 区分明显）
MEM_CARD_BG = {"≤2GB": "#e1f5fe", "3GB": "#b3e5fc", "4GB+": "#81d4fa", "待补": "#f6f7f9"}

# 屏幕显示技术徽章：卡片内显示的短标 + 是否高亮（进阶技术高亮，普通液晶/LED 弱化）
TECH_LABEL = {
    "LCD": "液晶", "普通液晶": "液晶",
    "LED": "LED", "LED(直下式)": "LED",
    "Mini LED": "MiniLED", "QD-Mini LED": "QD-Mini", "SQD-Mini LED": "SQD-Mini",
    "RGB-Mini LED": "RGB-Mini", "G+Mini LED": "G-Mini", "BGB-Mini LED": "BGB-Mini",
    "QLED": "QLED", "QD-LED": "QD-LED", "OLED": "OLED",
    "普通液晶(量子点)": "量子点", "LASER": "激光",
}
TECH_STRONG = [
    "Mini LED", "QD-Mini LED", "SQD-Mini LED", "RGB-Mini LED", "G+Mini LED", "BGB-Mini LED",
    "QLED", "QD-LED", "OLED", "普通液晶(量子点)", "LASER",
]

# 取数窗口元信息（由 build_v3.py 写出；缺失时兜底 W35）
META_SRC = os.path.join(BASE, "meta_v3.json")
def _read_meta():
    # meta_v3.json 可能被 iCloud 离载(dataless)：python open() 会触发云端拉取而阻塞；
    # 改用 `cat` 子进程（对 dataless 立即返回 Operation canceled），失败则用默认取数窗口。
    import subprocess
    try:
        p = subprocess.run(["cat", META_SRC], capture_output=True, text=True, timeout=15)
        if p.returncode == 0 and p.stdout.strip():
            return json.loads(p.stdout)
    except Exception:
        pass
    return {"week_end": 37, "week_start": 18, "fig1_window": "近20周(26W18-W37)", "fig1_gate": {"main": 1000, "big98": 200}}
META = _read_meta()

models = json.load(open(SRC, encoding="utf-8"))
# 剔除不展示型号（全部激光电视 + 指定款），数据保留但不再进入任何图
models = [m for m in models if not m.get("exclude_report")]
for i, m in enumerate(models):
    m["id"] = i
print(f"载入 {len(models)} 款型号  (AVC 26W01-26W{META['week_end']:02d} / 图1 窗口 {META['fig1_window']})")

# ============ 拼音首字母（搜索用，覆盖数据中出现的全部汉字） ============
PY = {
    "一":"y","东":"d","中":"z","京":"j","他":"t","代":"d","信":"x","光":"g",
    "入":"r","其":"q","分":"f","列":"l","创":"c","区":"q","发":"f","墨":"m",
    "壁":"b","大":"d","家":"j","寸":"c","小":"x","屏":"p","开":"k","心":"x",
    "新":"x","旗":"q","晶":"j","款":"k","海":"h","渠":"q","激":"j","炮":"p",
    "版":"b","玩":"w","现":"x","电":"d","端":"d","米":"m","系":"x","红":"h",
    "纸":"z","维":"w","能":"n","舰":"j","节":"j","视":"s","贴":"t","道":"d",
    "酷":"k","钢":"g","门":"m","随":"s","雀":"q","雷":"l","青":"q","高":"g",
    "鸟":"n","鹏":"p","鹤":"h",
}
def py_initials(s):
    """汉字取拼音首字母，ASCII 字母数字保留小写，其余跳过"""
    out = []
    for ch in s or "":
        if ch in PY:
            out.append(PY[ch])
        elif ch.isascii() and ch.isalnum():
            out.append(ch.lower())
    return "".join(out)

for m in models:
    # py = 品牌/系列/产品名 三段首字母缩写（空格分段），JS 侧供拼音搜索匹配
    m["py"] = " ".join(x for x in (
        py_initials(m.get("brand")), py_initials(m.get("series")), py_initials(m.get("product_name")),
    ) if x)

# ---- 图5 增长信号：读 fig5_growth.json（若缺失则不启用增长列） ----
_GRO = {}
_gp = os.path.join(BASE, "fig5_growth.json")
if os.path.exists(_gp):
    _GRO = json.load(open(_gp, encoding="utf-8"))
for m in models:
    _gd = _GRO.get(f"{m['brand']}||{m['model']}")
    m["gro"] = {
        "w8": (_gd or {}).get("w8") or [0]*8,
        "v26": (_gd or {}).get("v26") if _gd else m["cum_vol"],
        "first_num": (_gd or {}).get("first_num"),   # 首次出现周序号，如 202636
        "first_week": (_gd or {}).get("first_week"),
    }

# ==== 新品保护期：曾达标新品持久化 + 每款 nguard（绿点） ====
# 规则：达标新品(新品期内过 newGate 门槛) 摘标后仍保护 +16周（首现后满20周失效）。
# 保护期内豁免：清库、近8周全0、各图门槛；期间以绿点"新品保护期"标识（新品期仍显示「新品」chip）。
_WE_N = META['week_end']
_E = 202600 + _WE_N
_S4 = _E - 3
_GS = _E - 19                        # 首现后20周内皆保护（4周新+16周拓展）
def _n4(w): return sum((w or [])[-4:])
_flagd = False
for m in models:
    _fn = (m.get('gro') or {}).get('first_num')
    if _fn is not None and _S4 <= _fn <= _E:       # 当期新品窗口
        _z = m.get('size', 0)
        if _n4((m.get('gro') or {}).get('w8')) > (10 if _z >= 93 else 50):
            if not m.get('new_qualified'):
                m['new_qualified'] = True; _flagd = True
if _flagd:                                          # 持久化曾达标标记到源
    _src = json.load(open(SRC, encoding='utf-8'))
    _slist = _src if isinstance(_src, list) else _src.get('models', [])
    _kmap = {}
    for _x in _slist: _kmap['%s||%s' % (_x.get('brand'), _x.get('model'))] = _x
    for m in models:
        if m.get('new_qualified'):
            _k = '%s||%s' % (m['brand'], m['model'])
            if _k in _kmap: _kmap[_k]['new_qualified'] = True
    json.dump(_src, open(SRC, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
for m in models:
    _fn = (m.get('gro') or {}).get('first_num')
    m['nguard'] = bool(m.get('new_qualified')) and _fn is not None and _GS <= _fn <= _E

DATA = {
    "models": models,
    "brands": BRANDS, "color": COLOR,
    "useProduct": list(BRANDS_USE_PRODUCT),  # 4 个品牌展示产品名
    "fig1Cols": FIG1_COLS, "fig1Lab": {str(k): v for k, v in FIG1_LAB.items()},
    "sizeOrder": SIZE_ORDER, "beltOrder": BELT_ORDER,
    "partOrder": PART_ORDER,
    "memOrder": MEM_ORDER, "resOrder": RES_ORDER, "energyOrder": ENERGY_ORDER,
    "memCardBg": MEM_CARD_BG,  # 内存色：单卡底色（淡灰，避免视觉负担）
    "techLabel": TECH_LABEL, "techStrong": TECH_STRONG,  # 屏幕技术打标
    "meta": META,               # 取数窗口（week_end / fig1_window / fig1_gate）
}

# ============ CSS ============
CSS = """
*{margin:0;padding:0;box-sizing:border-box;}
:root{--labw:168px;--mc:130px;--mxn:8;}
body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:#f5f6f8;color:#1f2328;padding:18px;line-height:1.5;}
.wrap{max-width:1560px;margin:0 auto;}
header{background:#fff;border-radius:10px;padding:18px 22px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,.06);}
h1{font-size:20px;font-weight:700;margin-bottom:6px;}
.sub{font-size:12.5px;color:#6b7280;}
.sub b{color:#374151;}
.ctrl{background:#fff;border-radius:10px;padding:14px 18px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,.06);}
.crow{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:7px 0;border-bottom:1px dashed #eceff3;}
.crow:last-child{border-bottom:none;}
.clab{font-size:12px;color:#6b7280;min-width:52px;font-weight:600;flex-shrink:0;}
.chip{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border:1px solid #d8dde5;border-radius:15px;background:#fff;font-size:12.5px;cursor:pointer;user-select:none;transition:.12s;white-space:nowrap;}
.chip:hover{border-color:#9aa5b4;}
.chip.on{color:#fff;font-weight:600;}
.chip .dot{width:8px;height:8px;border-radius:50%;background:#c9cfd8;flex-shrink:0;}
.chip.on .dot{background:#fff;}
.btn{padding:5px 12px;border:1px solid #d8dde5;border-radius:6px;background:#fff;font-size:12.5px;cursor:pointer;transition:.12s;white-space:nowrap;}
.btn:hover{border-color:#9aa5b4;background:#f7f9fb;}
.btn.on{background:#1f2328;color:#fff;border-color:#1f2328;font-weight:600;}
.btn.mini{padding:3px 9px;font-size:11.5px;}
input.search{padding:6px 11px;border:1px solid #d8dde5;border-radius:6px;font-size:12.5px;width:230px;outline:none;}
input.search:focus{border-color:#0E7CE8;}
select{padding:6px 9px;border:1px solid #d8dde5;border-radius:6px;font-size:12.5px;background:#fff;cursor:pointer;outline:none;}
.sum{font-size:12px;color:#6b7280;margin-left:auto;}
.sum b{color:#0E7CE8;font-size:13px;}
.drop{position:relative;display:inline-block;}
.drop-panel{position:absolute;top:calc(100% + 5px);left:0;background:#fff;border:1px solid #d8dde5;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,.13);padding:8px;z-index:60;min-width:186px;max-width:calc(100vw - 36px);max-height:60vh;overflow:auto;display:none;}
.drop-panel.show{display:block;}
.drop-item{display:flex;align-items:center;gap:7px;padding:5px 8px;font-size:12.5px;cursor:pointer;border-radius:5px;white-space:nowrap;}
.drop-item:hover{background:#f2f5f9;}
.drop-item input{cursor:pointer;}
.drop-foot{display:flex;gap:6px;padding:6px 8px 2px;border-top:1px solid #eceff3;margin-top:5px;}
section{background:#fff;border-radius:10px;padding:16px 18px 18px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,.06);}
h2{font-size:15.5px;font-weight:700;margin-bottom:3px;display:flex;align-items:center;gap:8px;}
h2 .tag{font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:#eef4ff;color:#0E7CE8;}
.note{font-size:11.5px;color:#8b95a3;margin-bottom:11px;}
.scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;}
table{border-collapse:separate;border-spacing:2px;font-size:11.5px;}
table.ft{width:100%;table-layout:fixed;min-width:calc(var(--labw) + var(--mxn) * var(--mc));}
table.ft col.lab{width:var(--labw);}
table.ft col.d{width:calc((100% - var(--labw)) / var(--nc));}
th{padding:6px 8px;font-size:11px;color:#6b7280;font-weight:600;text-align:center;white-space:nowrap;background:#fafbfc;border-radius:4px;overflow:hidden;text-overflow:ellipsis;}
td{padding:0;vertical-align:top;}
.ft th:first-child,.ft td:first-child{position:sticky;left:0;z-index:2;}
.ft th:first-child{background:#fafbfc;box-shadow:3px 0 0 #fafbfc;}
.ft td:first-child{background:#fff;box-shadow:3px 0 0 #fff,6px 0 8px -4px rgba(15,23,42,.14);}
.rl{display:flex;align-items:center;gap:6px;padding:4px 9px 4px 0;font-size:12px;white-space:nowrap;min-width:0;overflow:hidden;}
.rl .bar{width:3px;height:15px;border-radius:2px;flex-shrink:0;}
.rl .bn{color:#8b95a3;font-size:10.5px;flex-shrink:0;}
.rl .sn{font-weight:600;color:#1f2328;overflow:hidden;text-overflow:ellipsis;}
.rl .cnt{color:#9aa5b4;font-size:10.5px;margin-left:auto;flex-shrink:0;}
.cell{width:100%;min-width:0;height:34px;border-radius:5px;display:flex;flex-direction:column;align-items:center;justify-content:center;cursor:pointer;transition:.1s;border:1px solid transparent;position:relative;overflow:hidden;}
.cell:hover{border-color:#0E7CE8;}
.cell .v{font-size:12.5px;font-weight:600;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cell .n{font-size:9.5px;opacity:.72;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
/* 图2 格子内「销量 + 款数」并排一行（替代原两行，避免与底部堆叠条重叠） */
.cell.hasbar .vn{display:inline-flex;align-items:baseline;gap:3px;max-width:100%;min-width:0;}
.cell.empty{background:#fafbfc;cursor:default;}
.cell.empty:hover{border-color:transparent;}
.cell.zero{background:#fafbfc;color:#c9cfd8;cursor:default;}
.cell.zero:hover{border-color:transparent;}
/* 图2 方案 A：格内品牌份额堆叠条 */
.cell .sbar{position:absolute;left:3px;right:3px;bottom:3px;height:7px;display:flex;border-radius:2px;overflow:hidden;pointer-events:none;}
.cell .sbar i{display:block;height:100%;}
.cell.hasbar{padding-bottom:9px;}
/* 图2 展开面板品牌筛选 chips */
.exp-brands{display:flex;flex-wrap:wrap;gap:4px;margin:2px 0 8px;}
.ebtn{font-size:10.5px;line-height:1.4;padding:2px 9px;border-radius:4px;color:#fff;cursor:pointer;user-select:none;transition:.12s;}
.ebtn.off{opacity:.3;}
.gridbox{display:flex;flex-wrap:wrap;gap:4px;padding:6px;min-height:38px;align-items:flex-start;}
.mc{display:inline-flex;align-items:stretch;border-radius:4px;overflow:hidden;font-size:10.5px;cursor:pointer;border:1px solid #e3e8ef;background:#fff;width:100%;max-width:168px;min-width:0;transition:.1s;}
.mc.memna{border-style:dashed;border-color:#c9cfd8;}
.mc:hover{border-color:#0E7CE8;box-shadow:0 1px 4px rgba(14,124,232,.22);}
.mc.sel{border-color:#0E7CE8;border-width:2px;box-shadow:0 0 0 2px rgba(14,124,232,.18);}
.mc .mbar{width:3px;flex-shrink:0;}
.mc .mbody{padding:3px 6px;min-width:0;flex:1 1 auto;}
.mc .mtop{display:flex;align-items:center;gap:4px;}
.mc .md{width:6px;height:6px;border-radius:50%;flex-shrink:0;}
.mc .mnm{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1 1 auto;min-width:0;}
.mc .mbot{font-size:9.5px;color:#8b95a3;display:flex;gap:4px;align-items:center;white-space:nowrap;overflow:hidden;}
.mc .badge{font-size:8.5px;padding:0 3px;border-radius:2px;background:#000;color:#fff;font-weight:600;flex-shrink:0;}
.mc .badge.tech{background:#f0f2f5;color:#8b95a3;font-weight:500;white-space:nowrap;flex:0 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;max-width:64px;}
.mc .badge.tech.hi{background:#0E7CE8;color:#fff;font-weight:600;}
.mc .badge.tech.none{background:#fafbfc;color:#c9cfd8;border:1px dashed #dde2e8;padding:0 2px;}
.mc .badge.alt{background:#f5f7fa;color:#98a2b0;font-weight:400;}
.mc .badge.clearb{background:#e11d48;color:#fff;font-weight:600;border-radius:2px;font-size:8px;padding:0 3px;line-height:1.4;}
.cmpOn .mc{cursor:crosshair;}

/* ===== 图3/图4 参数筛选按钮（行头 + 列表头，点击多选过滤本图） ===== */
.ftools{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 8px;}
.fbtn{display:inline-flex;align-items:center;justify-content:center;gap:4px;
  padding:2px 7px;border:1px solid #dde2e8;border-radius:4px;background:#fff;
  font-size:10.5px;line-height:1.5;color:#5b6675;cursor:pointer;user-select:none;
  transition:border-color .12s,background .12s,opacity .12s;white-space:nowrap;}
.fbtn:hover{border-color:#8fb6e8;background:#f5f9ff;}
.fbtn.on{border-color:#0E7CE8;background:#eef4ff;color:#0E7CE8;font-weight:600;}
.fbtn.off{opacity:.5;background:#fafbfc;color:#98a2b0;}
.fbtn.all{border-color:#0E7CE8;color:#0E7CE8;}
.fbtn.all.off{opacity:.55;border-color:#dde2e8;color:#98a2b0;background:#fafbfc;}
th .fbtn.col{display:inline-block;text-align:center;padding:3px 6px;font-size:11px;font-weight:600;}
th .fbtn.col .cnt-label{display:block;font-size:9px;font-weight:400;opacity:.75;margin-top:1px;}
.fsum{font-size:10.5px;color:#8b95a3;}
.fsum b{color:#0E7CE8;font-weight:600;}
/* 每图独立的品牌多选条 */
.fig-brand{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin:0 0 8px;}
.fig-brand .clab{min-width:auto;margin-right:2px;font-size:12px;color:#6b7280;font-weight:600;}
.fig-brand .chip{padding:4px 10px;font-size:12px;}
.fig-brand .slab{font-size:12px;color:#6b7280;font-weight:600;margin-right:2px;}
.fig-brand .chip.on:not([style]){background:#0E7CE8;border-color:#0E7CE8;}
/* 各图「移除清库机」开关条 */
.clrctl{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 8px;padding:6px 10px;background:#fff8f5;border:1px dashed #f5a97f;border-radius:6px;font-size:11.5px;color:#8a4b2d;}
.clrctl .ctl-label{display:inline-block;font-size:10px;font-weight:700;color:#fff;background:#f97316;border-radius:3px;padding:1px 6px;}
.clrbtn{display:inline-flex;align-items:center;gap:5px;padding:4px 11px;border:1px solid #f5a97f;background:#fff;color:#c2410c;border-radius:14px;cursor:pointer;font-size:12px;font-weight:600;transition:.12s;white-space:nowrap;}
.clrbtn:hover{border-color:#c2410c;background:#fff3ec;}
.clrbtn.on{background:#c2410c;border-color:#c2410c;color:#fff;}
.clrbtn.on:hover{opacity:.9;}
.clr-def{color:#a35c35;font-size:10.5px;line-height:1.5;flex:1 1 260px;}
.clr-def b{color:#7c3f1f;font-weight:700;}
tr.frow.off td:not(.rowhead){opacity:.26;}
tr.frow.off .mc{pointer-events:none;}
td.col-off{background:#fafbfc;border-radius:5px;min-height:38px;}
.empty-dim{font-size:11px;color:#ccd3dc;padding:6px 0;display:inline-block;}
/* ===== 图3/图4 品牌手风琴（方案A） ===== */
.blist{display:flex;flex-direction:column;gap:8px;}
.bsec{border:1px solid #e3e8ef;border-radius:8px;overflow:hidden;background:#fff;}
.bhead{display:flex;align-items:center;gap:8px;padding:7px 10px;cursor:pointer;background:#f7f9fb;border-bottom:1px solid #eef1f6;flex-wrap:wrap;}
.bhead .tri{color:#0E7CE8;font-size:11px;width:14px;flex-shrink:0;transition:transform .12s;}
.bhead .tri.on{transform:rotate(90deg);}
.bn-b{font-weight:700;color:#5b6675;padding-left:6px;}
.brep{font-size:10.5px;color:#0E7CE8;background:#eef4ff;border:1px solid #b9d3f5;border-radius:11px;padding:1px 8px;cursor:help;white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis;}
.bcnt{margin-left:auto;font-size:11px;color:#8b95a3;flex-shrink:0;}
.bsub{padding:8px;background:#fff;}
.bsub-tb{border-collapse:collapse;width:100%;font-size:11.5px;}
.bsub-tb th,.bsub-tb td{border:1px solid #eef1f6;padding:5px 6px;text-align:center;vertical-align:top;}
.bsub-tb th{background:#f7f9fb;color:#6b7280;font-weight:600;white-space:nowrap;}
.bsub-tb td.blab{background:#fafbfc;font-weight:600;color:#4b5563;white-space:nowrap;text-align:left;width:86px;}
.bsub-tb .gridbox{min-height:32px;padding:3px;}
/* ===== 二轴网格单元格内：同品牌折叠组 ===== */
.bgg{flex:1 1 100%;min-width:0;margin:0 0 3px;}
.bgg:last-child{margin-bottom:0;}
.bghead{display:flex;align-items:center;gap:6px;padding:2px 6px;cursor:pointer;background:#f7f9fb;border:1px solid #edf0f5;border-left:4px solid #888;border-radius:5px;font-size:11px;}
.bghead:hover{background:#eef4ff;}
.bghead .gtri{color:#0E7CE8;font-size:10px;width:12px;flex-shrink:0;transition:transform .12s;}
.bghead .gtri.on{transform:rotate(90deg);}
.gbn{font-weight:700;color:#5b6675;white-space:nowrap;}
.grep{font-size:10px;color:#0E7CE8;background:#eef4ff;border:1px solid #b9d3f5;border-radius:9px;padding:0 6px;cursor:help;white-space:nowrap;max-width:180px;overflow:hidden;text-overflow:ellipsis;}
.gcnt{margin-left:auto;color:#8b95a3;font-size:10px;flex-shrink:0;}
.bgbody{padding:5px 2px 2px;}
.gridbox-in{display:flex;flex-wrap:wrap;gap:4px;align-items:flex-start;}
.exp{background:#f7f9fb;border-radius:6px;padding:8px;margin-top:6px;}
.exp-h{font-size:11.5px;color:#6b7280;margin-bottom:5px;font-weight:600;}
.exp-fold{margin-left:8px;color:#0E7CE8;font-weight:600;cursor:pointer;user-select:none;font-size:11px;}
.exp-fold:hover{text-decoration:underline;}
#cmpBar{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:2px solid #0E7CE8;box-shadow:0 -4px 18px rgba(0,0,0,.13);padding:10px 20px;display:none;z-index:80;}
#cmpBar.show{display:block;}
.cb-in{max-width:1560px;margin:0 auto;display:flex;align-items:center;gap:9px;flex-wrap:wrap;}
.cb-t{font-size:12.5px;font-weight:700;flex-shrink:0;}
.cb-list{display:flex;gap:6px;flex-wrap:wrap;flex:1;}
.cb-i{display:inline-flex;align-items:center;gap:5px;padding:3px 8px;border:1px solid #d8dde5;border-left-width:3px;border-radius:5px;font-size:11.5px;background:#fff;}
.cb-i .x{cursor:pointer;color:#9aa5b4;font-weight:700;padding:0 2px;}
.cb-i .x:hover{color:#ED1C24;}
.modal{position:fixed;inset:0;background:rgba(15,23,42,.5);display:none;z-index:100;align-items:center;justify-content:center;padding:22px;}
.modal.show{display:flex;}
.mbox{background:#fff;border-radius:11px;max-width:1180px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 20px 60px rgba(0,0,0,.28);}
.mhd{display:flex;align-items:center;gap:10px;padding:14px 18px;border-bottom:1px solid #eceff3;position:sticky;top:0;background:#fff;z-index:2;}
.mhd h3{font-size:15px;font-weight:700;}
.mhd .x{margin-left:auto;cursor:pointer;font-size:21px;color:#9aa5b4;line-height:1;}
.mhd .x:hover{color:#1f2328;}
.mbd{padding:14px 18px 20px;}
.ctab{border-collapse:collapse;width:100%;font-size:12px;min-width:480px;}
.ctab th,.ctab td{border:1px solid #eceff3;padding:7px 10px;text-align:center;vertical-align:middle;}
.ctab th{background:#f7f9fb;font-weight:600;color:#6b7280;font-size:11.5px;white-space:nowrap;}
.ctab th:first-child,.ctab td.lab{position:sticky;left:0;z-index:2;}
.ctab th:first-child{background:#f7f9fb;}
.ctab td.lab{background:#fafbfc;font-weight:600;color:#4b5563;white-space:nowrap;}
.ctab td.best{background:#e8f7ee;color:#0f7b3d;font-weight:600;}
.ctab td.worst{background:#fdecec;color:#c02929;font-weight:600;}
.ctab td.na{color:#c2c8d0;font-style:italic;}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:11px;color:#6b7280;margin-top:10px;padding-top:9px;border-top:1px dashed #eceff3;align-items:center;}
.mem-legend{border-top:none;margin-top:9px;padding-top:0;}
.lg{display:flex;align-items:center;gap:5px;}
.lg i{width:11px;height:11px;border-radius:3px;display:inline-block;}
.lg i.circ{border-radius:50%;width:9px;height:9px;}
/* 图1 热力刻度图例：浅→深 = 销量强度 */
.heat-scale{display:inline-flex;align-items:center;gap:5px;}
.heat-scale .hs-bar{width:64px;height:9px;border-radius:3px;background:linear-gradient(90deg,#f2f6f9,#7fb3e8,#0E7CE8);border:1px solid #dde2e8;}
footer{text-align:center;color:#9aa5b4;font-size:11.5px;padding:14px 0 8px;}
.warn{background:#fff8e6;border:1px solid #ffe1a6;border-radius:8px;padding:10px 14px;margin-bottom:14px;font-size:12px;color:#7a5600;}
.warn b{color:#5c4000;}

/* ===== 四图表头浮动吸顶（仿参考报告：克隆 thead，不压缩表格高度） ===== */
.float-head{position:fixed;top:0;left:0;z-index:900;overflow:hidden;display:none;
  background:#fff;box-shadow:0 2px 8px rgba(0,0,0,.15);pointer-events:none;}
.float-head table.ft th{background:#f7f9fb;}
.float-head table.ft th:first-child{background:#f7f9fb;box-shadow:3px 0 0 #f7f9fb;}
.float-head table.ft{margin:0;}
.float-head table.f5tbl th{background:#f7f9fb;white-space:nowrap;font-size:11.5px;font-weight:600;color:#6b7280;}
.float-head table.f5tbl th:first-child{box-shadow:3px 0 0 #f7f9fb;}
.float-head table.f5tbl{margin:0;}

/* ===== 图1 单元格明细（仿参考报告 .m-row / .m-count / .m-vol） ===== */
.cell.f1{padding:5px 6px;height:auto;min-height:38px;align-items:stretch;justify-content:flex-start;overflow:visible;}
.cell.f1 .m-count{font-size:9.5px;opacity:.72;font-weight:400;line-height:1.3;}
.cell.f1 .m-list{display:flex;flex-direction:column;gap:1px;width:100%;min-width:0;}
.cell.f1 .m-row{display:flex;align-items:baseline;gap:4px;font-size:10.5px;line-height:1.35;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0;}
.cell.f1 .m-name{font-weight:600;overflow:hidden;text-overflow:ellipsis;min-width:0;}
.cell.f1 .m-vol{font-size:9.5px;opacity:.75;flex-shrink:0;margin-left:auto;}
.cell.f1 .m-clr{font-style:normal;font-size:8px;font-weight:700;color:#e11d48;background:rgba(255,255,255,.85);border-radius:2px;padding:0 3px;line-height:1.4;flex-shrink:0;margin-left:2px;}
.cell.f1 .m-more{font-size:9px;opacity:.6;font-weight:400;}
/* 行头品牌色左边条（仿参考报告 4px border-left） */
.rl{border-left:4px solid transparent;padding-left:7px;}
.f1-empty{background:#fcfcfc;min-height:38px;}

@media (max-width:768px){
  .cell.f1{padding:4px 5px;min-height:34px;}
  .cell.f1 .m-row{font-size:10px;}
  .cell.f1 .m-vol{font-size:9px;}
}

@media (max-width:1024px){
  body{padding:12px;}
  .wrap{max-width:100%;}
  header{padding:15px 16px;}
  h1{font-size:18px;}
  .sub{font-size:12px;line-height:1.65;}
  .ctrl{padding:11px 13px;}
  section{padding:13px 13px 15px;}
  input.search{width:190px;}
}
@media (max-width:768px){
  :root{--labw:132px;--mc:118px;}
  body{padding:9px;line-height:1.45;}
  header{padding:13px 13px;}
  h1{font-size:16.5px;}
  .ctrl{padding:10px 11px;}
  .crow{gap:7px;padding:6px 0;}
  .clab{min-width:auto;width:100%;}
  .sum{margin-left:0;width:100%;}
  input.search{width:100%;flex:1 1 100%;}
  h2{font-size:14px;flex-wrap:wrap;gap:6px;}
  h2 .tag{font-size:10px;}
  .note{font-size:11px;margin-bottom:9px;}
  section{padding:12px 11px 14px;margin-bottom:11px;}
  .legend{gap:9px;font-size:10.5px;}
  .ft td:first-child{box-shadow:3px 0 0 #fff,5px 0 7px -3px rgba(15,23,42,.16);}
  .modal{padding:10px;align-items:flex-start;}
  .mbox{max-height:94vh;border-radius:9px;}
  .mhd{padding:11px 13px;gap:8px;}
  .mhd h3{font-size:14px;}
  .mbd{padding:11px 12px 16px;}
  .ctab{font-size:11.5px;}
  .ctab th,.ctab td{padding:6px 7px;}
  #cmpBar{padding:8px 11px;}
  .cb-in{gap:7px;}
  .cb-t{width:100%;}
  .cb-i{font-size:11px;}
  footer{font-size:11px;padding:12px 0 6px;}
}
@media (max-width:480px){
  :root{--labw:112px;--mc:106px;}
  body{padding:7px;}
  .chip{padding:4px 9px;font-size:11.5px;}
  .btn{padding:4px 9px;font-size:11.5px;}
  .btn.mini{padding:3px 8px;font-size:11px;}
  h1{font-size:15.5px;}
  .wrap>section>.note{font-size:10.5px;}
  .cell{height:31px;}
  .cell .v{font-size:11.5px;}
  .mc{font-size:10px;}
  .mc .badge.tech{max-width:52px;}
  .fbtn{padding:2px 5px;font-size:10px;}
  th .fbtn.col{padding:2px 4px;font-size:10px;}
  .fsum{font-size:10px;}
  .modal{padding:6px;}
  .mbox{max-height:96vh;}
}
@media (hover:none){
  .cell:hover{border-color:transparent;}
  .mc:hover{border-color:#e3e8ef;box-shadow:none;}
  .chip:hover{border-color:#d8dde5;}
  .btn:hover{border-color:#d8dde5;background:#fff;}
}

/* ===== 图5 单产品定位分析 ===== */
.f5wrap{display:flex;gap:14px;align-items:stretch}
.f5left{flex:0 0 216px;display:flex;flex-direction:column;gap:10px}
.f5card{background:#f7f9fb;border:1px solid #e3e8ef;border-radius:9px;padding:10px}
.f5card h4{margin:0 0 8px;font-size:12px;font-weight:700;color:#1f2328}
.f5field{margin-bottom:7px}
.f5field label{display:block;font-size:10px;color:#6b7280;margin-bottom:3px}
.f5seg{display:flex;flex-wrap:wrap;gap:4px}
.f5seg span{flex:1;min-width:52px;text-align:center;font-size:11px;padding:4px 2px;border:1px solid #d8dde5;border-radius:6px;cursor:pointer;color:#6b7280;background:#fff;white-space:nowrap}
.f5seg span:hover{border-color:#9aa5b4}
.f5seg span.on{color:#fff;background:#0E7CE8;border-color:#0E7CE8;font-weight:600}
/* 图5 价位段12档：窄栏内换行 3 列，避免文本溢出格子 */
#f5pr.f5seg span{flex:1 1 31%;min-width:0;font-size:9.5px;padding:4px 1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;letter-spacing:-.4px;border-radius:4px}
.f5badge{border-radius:8px;padding:8px 10px;font-size:11px;line-height:1.5}
.f5badge .bt{font-weight:700;font-size:12px;margin-bottom:3px}
.f5badge.opp{border:1px solid #f59e0b;background:#fff7e6}
.f5badge.opp .bt{color:#b45309}
.f5badge.grow{border:1px solid #2563eb;background:#eef4ff}
.f5badge.grow .bt{color:#1d4ed8}
.f5badge.gap{border:1px solid #16a34a;background:#eafaf0}
.f5badge.gap .bt{color:#15803d}
.f5badge.mid{border:1px solid #94a3b8;background:#f1f5f9}
.f5badge.mid .bt{color:#475569}
/* 对标表列筛选条 */
.f5filterbar{display:none;align-items:center;gap:8px;flex-wrap:wrap;margin:6px 0;padding:6px 8px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px}
.f5filterbar.show{display:flex}
.f5filterbar .f5fbt{font-weight:700;font-size:11px;color:#0E7CE8;white-space:nowrap}
.f5filterbar .f5fchips{display:flex;flex-wrap:wrap;gap:4px;flex:1}
.f5filterbar .f5fact{display:flex;gap:4px;white-space:nowrap}
/* 表头筛选按钮 & 相对档位徽章 */
th .f5fbtn{padding:0 4px;margin-right:2px;font-size:10px;border:1px solid #d8dde5;border-radius:4px;background:#fff;color:#0E7CE8;cursor:pointer;vertical-align:middle}
th .f5fbtn:hover{background:#0E7CE8;color:#fff}
.f5rel{display:inline-block;margin-left:4px;font-size:10px;font-weight:700;border-radius:3px;padding:0 3px;vertical-align:middle}
.f5rel.win{background:#dcfce7;color:#15803d}
.f5rel.tie{background:#f1f5f9;color:#64748b}
.f5rel.loss{background:#fee2e2;color:#b91c1c}
/* 近8周趋势微型柱状图（修复：<i> 需 block/inline-flex 才可见） */
.f5spark{display:inline-flex;align-items:flex-end;gap:2px;height:88px;vertical-align:middle;line-height:0}
.f5spark i{display:inline-block;width:5px;min-height:2px;border-radius:1.5px;background:#38bdf8}
.f5spark.flat i{background:#e2e8f0}
.f5clr{display:block;color:#dc2626;font-size:10px;font-weight:700;line-height:1.15;margin-bottom:3px;letter-spacing:1px}
/* 表头 ▾ 列下拉筛选面板（fixed 定位避免被 .scroll 裁剪） */
.f5ddown{display:none;position:fixed;z-index:9999;min-width:180px;max-width:240px;max-height:360px;overflow:auto;background:#fff;border:1px solid #d8dde5;border-radius:9px;box-shadow:0 10px 30px rgba(15,23,42,.14)}
.f5ddown.show{display:block}
.f5ddown .f5dd-t{display:flex;align-items:center;justify-content:space-between;padding:7px 10px;font-size:11px;font-weight:700;color:#1f2328;border-bottom:1px solid #eef1f5;position:sticky;top:0;background:#fff;z-index:1}
.f5ddown .f5dd-x{cursor:pointer;color:#64748b;font-weight:400}
.f5ddown .f5dd-list{padding:5px;max-height:250px;overflow:auto}
.f5ddown .f5dd-it{display:flex;align-items:center;gap:7px;padding:5px 7px;border-radius:5px;font-size:12px;color:#1f2328;cursor:pointer}
.f5ddown .f5dd-it:hover{background:#f1f5f9}
.f5ddown .f5dd-it.on{background:#e0f2fe}
.f5ddown .f5dd-it input{margin:0}
.f5ddown .f5dd-it em{margin-left:auto;font-style:normal;font-size:10px;color:#94a3b8}
.f5ddown .f5dd-f{display:flex;gap:6px;padding:7px 9px;border-top:1px solid #eef1f5}
.f5ddown .f5dd-f button{font-size:11px;padding:3px 9px;border:1px solid #d8dde5;border-radius:5px;background:#fff;color:#1f2328;cursor:pointer}
.f5ddown .f5dd-f button:hover{border-color:#0E7CE8;color:#0E7CE8}
th .f5fbtn.act{background:#0E7CE8;color:#fff;border-color:#0E7CE8}
.f5main{flex:1;min-width:0}
.f5maptitle{font-size:11.5px;color:#6b7280;margin-bottom:6px}
.f5map{border:1px solid #e3e8ef;border-radius:9px;overflow:hidden;background:#fff}
.f5gh{display:flex;margin-left:46px;border-bottom:1px solid #eef0f5}
.f5gh span{flex:1;text-align:center;font-size:10px;color:#6b7280;padding:6px 2px;white-space:nowrap}
.f5row{display:flex;border-bottom:1px solid #f4f6f9}
.f5row:last-child{border-bottom:none}
.f5ylab{flex:0 0 46px;display:flex;align-items:center;padding-left:8px;font-size:11px;color:#6b7280;border-right:1px solid #eef0f5;font-weight:600}
.f5cells{display:flex;flex:1}
.f5cell{flex:1;min-width:0;aspect-ratio:1.35;margin:2px;border-radius:5px;font-size:10px;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#1f2328;position:relative;border:1px solid transparent;cursor:pointer;user-select:none;transition:transform .1s;overflow:hidden}
.f5cell:hover{transform:scale(1.06);z-index:3;box-shadow:0 2px 6px rgba(15,23,42,.14)}
.f5cell b{font-size:12px;line-height:1.2}
.f5cell small{color:#6b7280}
.f5cell.gap{border-style:dashed}
.f5cell.sel{outline:2px solid #ef4444;outline-offset:1px}
.f5cell .gg{position:absolute;top:2px;right:4px;color:#fff;font-size:8.5px;padding:0 4px;border-radius:3px}
.f5cell .gg.y{background:#2563eb}
.f5xlab{margin-left:46px;display:flex}
.f5xlab span{flex:1;text-align:center;font-size:10px;color:#6b7280;padding:6px 2px;white-space:nowrap}
.f5legend{display:flex;gap:12px;margin-top:9px;font-size:10.5px;color:#6b7280;flex-wrap:wrap;align-items:center}
.f5legend .lg{display:flex;align-items:center;gap:4px}
.f5sw{width:15px;height:11px;border-radius:3px;display:inline-block;border:1px solid rgba(0,0,0,.06)}
.f5tbl{width:100%;border-collapse:collapse;font-size:11px;min-width:760px}
.f5tbl th,.f5tbl td{padding:5px 7px;border-bottom:1px solid #eef0f5;text-align:left;white-space:nowrap}
.f5tbl th{color:#6b7280;font-size:10px;font-weight:600;background:#fafbfc;cursor:pointer;user-select:none}
.f5tbl th:hover{background:#f1f5f9}
.f5tbl tr.me{background:#fef2f2}
.f5tbl .up{color:#16a34a;font-weight:700}
.f5tbl .down{color:#ef4444;font-weight:700}
.f5tbl .na{color:#9aa5b4}
.f5tier{display:inline-block;min-width:34px;font-size:10.5px;font-weight:700;border-radius:9px;padding:1px 7px;color:#fff}
.f5tier.hi{background:#1d4ed8}
.f5tier.md{background:#64748b}
.f5tier.lo{background:#f59e0b}
.f5score{margin-top:12px;padding:10px 12px;background:#f8fafc;border:1px solid #e6e9f0;border-radius:8px;font-size:10.5px;color:#475569;line-height:1.9}
.f5score b{color:#1e293b;margin-right:6px}
.f5score>div{display:flex;flex-wrap:wrap;gap:2px 18px;margin-top:2px}
.f5score span{white-space:nowrap}
.f5deep{margin-top:0;border:1px solid #e6e9f0;border-radius:10px;overflow-x:auto;overflow-y:hidden;font-size:11.5px;color:#334155}
.f5deep .dhead{background:#f1f5f9;padding:8px 12px;font-weight:700;color:#1e293b;font-size:12.5px;border-bottom:1px solid #e2e8f0}
.f5deep .dsec{padding:9px 12px;border-top:1px solid #eef0f5}
.f5deep .dsec .dt{font-size:11px;font-weight:700;color:#2563eb;margin-bottom:7px}
.f5deep .drow{display:flex;align-items:center;gap:6px;margin:3px 0;font-size:10.5px;line-height:14px}
.f5deep .dbar{flex:1 1 auto;min-width:14px;height:8px;border-radius:2px;background:linear-gradient(to right,#60a5fa var(--w,0%),#eef2f7 var(--w,0%))}
.f5deep .dbn{flex:0 0 30px;color:#475569;text-align:right;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
.f5deep .dbv{flex:0 0 82px;color:#64748b;white-space:nowrap;text-align:left}
.f5deep table.f5dtbl{width:100%;border-collapse:collapse;font-size:10.5px;margin-top:2px}
.f5deep table.f5dtbl th{background:#fafbfc;color:#6b7280;font-weight:600;padding:4px 7px;border:1px solid #e6e9f0;text-align:left;white-space:nowrap}
.f5deep table.f5dtbl td{padding:4px 7px;border:1px solid #eef0f5;white-space:nowrap}
.f5deep .dscroll{overflow-x:auto;margin-top:2px}
.f5deep .dt2{font-size:10.5px;font-weight:700;color:#374151;margin:10px 0 4px}
.f5deep .gm{font-size:11px;line-height:1.7;color:#475569}
.f5deep .gm.sum{margin-top:6px;color:#334155;background:#f1f5f9;padding:6px 8px;border-radius:6px}
.f5deep .gm.adv{color:#1d4ed8;padding-left:2px}
.f5deep .gchip{display:inline-block;background:#eff6ff;border:1px solid #bfdbfe;color:#1d4ed8;border-radius:12px;padding:1px 8px;margin:2px 4px 2px 0;font-size:10.5px}
.f5deep .gchip b{color:#dc2626;margin-left:3px}
.f5anchor{cursor:pointer}
.f5badge-new{color:#2563eb;font-size:9px}
.f5topline{display:flex;align-items:baseline;gap:8px;margin-bottom:4px;flex-wrap:wrap}
.f5topline .tcap{font-size:12px;font-weight:700}
.f5topline .tn{font-size:11.5px;color:#6b7280}

/* ===== 新品（近4周首现滚动） ===== */
.badge.ng{background:#10b981;color:#fff;font-weight:600}
.ngchip{display:inline-block;background:#10b981;color:#fff;border-radius:9px;font-size:9px;font-weight:700;padding:1px 6px;vertical-align:1px;margin-left:3px}
.m-ng{background:#d1fae5;color:#047857;font-weight:700;font-size:9px;padding:0 4px;border-radius:6px;margin-right:3px}
.gdot{display:inline-block;width:8px;height:8px;border-radius:50%;background:#10b981;vertical-align:1px;margin:0 3px;box-shadow:0 0 0 1px rgba(255,255,255,.7);cursor:help}
/* 新品上市板块 */
.newsec{border:2px solid #10b981;border-radius:12px;background:#f0fdf4;padding:16px 20px;margin-bottom:16px}
.newsec .hd{display:flex;align-items:center;gap:8px;margin-bottom:6px}
.newsec .hd h2{font-size:16px;color:#065f46;font-weight:700}
.newsec .hd .cnt{font-size:12px;color:#047857}
.newsec .sub{font-size:11px;color:#15803d;margin-bottom:12px}
.newgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:10px}
.newcard{background:#fff;border:1px solid #a7f3d0;border-radius:8px;padding:8px 10px}
.newcard .top{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.newcard .bn{font-weight:600;font-size:12px}
.newcard .tag-done{background:#10b981;color:#fff;border-radius:8px;font-size:9px;padding:0 5px}
.newcard .tag-watch{background:#f59e0b;color:#fff;border-radius:8px;font-size:9px;padding:0 5px}
.newcard .meta{font-size:10.5px;color:#6b7280;margin-top:3px}
.newcard .meta b{color:#334155}
.newcard .ncnt{font-style:normal;font-size:9px;background:#d1fae5;color:#047857;border-radius:6px;padding:0 4px;margin-left:2px}
.newcard .nlist{display:flex;flex-direction:column;gap:5px;margin-top:7px;border-top:1px dashed #a7f3d0;padding-top:6px}
.nitem{background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:5px 7px}
.nitem .itop{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.nitem .sz{font-weight:700;font-size:12px;color:#065f46}
.nitem .mrl{font-size:11px;color:#475569}
.nitem .imet{font-size:10px;color:#6b7280;margin-top:2px}
.nitem .imet b{color:#334155}
.nitem .ex{font-size:9px;color:#f59e0b}
"""

# ============ JS ============
JS = r"""
const D = window.__DATA__;
const M = D.models;
const CO = D.color;
// 页面级错误显示（便于定位对标表偶发不刷新）
function f5Err(e, ctx){
  try{ const el = document.getElementById('f5tbl');
    if(el) el.innerHTML = '<div class="na" style="color:#b91c1c;font:12px/1.5 monospace;white-space:pre-wrap;max-width:760px">图5渲染中断['+ctx+']：'+ (e&&e.message||e) +'\n'+ ((e&&e.stack||'').split('\n').slice(0,3).join('\n')) +'</div>';
  }catch(_){}
}
window.addEventListener('error', ev => { try{ const el=document.getElementById('f5tbl'); if(el) el.innerHTML='<div class="na" style="color:#b91c1c;font:12px monospace">JS错误：'+ev.message+' @'+(ev.lineno||'')+':'+(ev.colno||'')+'</div>'; }catch(_){} });
const BRANDS_USE_PRODUCT = new Set(D.useProduct || []);
const SZ3 = ['55吋','65吋','75吋','85吋','98吋+'];            // 图3 尺寸筛选档
const SZ4 = ['32吋','40吋','43吋','50吋'];                    // 图4 尺寸筛选档
function szOf3(m){ const s=m.size||0; return s<=60?'55吋':s<=69?'65吋':s<=79?'75吋':s<=92?'85吋':'98吋+'; }
function szOf4(m){ const s=m.size||0; return s<=35?'32吋':s<=41?'40吋':s<=46?'43吋':'50吋'; }
const S = {
  b: {1:new Set(D.brands), 2:new Set(D.brands), 3:new Set(D.brands), 4:new Set(D.brands)},  // 每图独立品牌集合
  sz3: new Set(SZ3),          // 图3 尺寸筛选集合
  sz4: new Set(SZ4),          // 图4 尺寸筛选集合
  size: '全部',
  belts: new Set(),
  q: '',
  mainOnly: false,
  cmp: false,
  ncols: [],
  sel: [],
  fig4: 'energy',
  fig3dim: 'refresh',        // 图3 子网格列维度：'refresh' | 'mem'
  f3g: new Set(),            // 图3 网格单元格内已展开的品牌组（key=行||列||品牌）
  f4g: new Set(),            // 图4 网格单元格内已展开的品牌组
  open: {},
  // 图3/图4 行/列筛选状态（每图独立：null=全选，Set=仅显示选中项）
  f3Rows: null, f3Cols: null,
  f4Rows: null, f4Cols: null,
  expBrands: {},            // 图2 展开面板的品牌筛选：{ 展开key: Set(品牌) | undefined }
  clr: {1:0,2:0,3:0,4:0,5:0},   // 各图「移除清库机」开关：0=显示清库(打标) 1=隐藏清库型号
};

const fmt = n => n >= 10000 ? (n/10000).toFixed(1)+'万' : String(n);
const beltOf = m => m.price_belt || '待补';

function searchHit(m) {   // 搜索框：品牌+系列+型号+产品名+拼音首字母；支持空格多词组合与品牌+系列连写/前缀（海信E7Q、海信E7、海信E）
  if (!S.q) return true;
  const hay = ((m.brand||'')+' '+(m.series||'')+' '+(m.model||'')+' '+(m.product_name||'')+' '+(m.py||'')).toLowerCase();
  const q0 = S.q.toLowerCase().trim();
  if (!q0) return true;
  const toks = q0.split(/\s+/).filter(Boolean);   // 空格拆词：每词都需命中拼接串（多词组合）
  const okTok = toks.length && toks.every(t => hay.includes(t));
  const qTight = q0.replace(/\s+/g,'');           // 去所有空格后的连写串
  const okTight = hay.replace(/\s+/g,'').includes(qTight);  // 品牌+系列 紧凑连写/前缀
  return okTok || okTight;
}
function pass(m) {   // 全局筛选（尺寸/价格带/搜索），不含品牌
  if (S.size !== '全部' && m.size_bucket !== S.size) return false;
  if (S.belts.size && !S.belts.has(beltOf(m))) return false;
  return searchHit(m);
}
function filtered() { return M.filter(pass); }
/* 近8周0销量排除（图1-4 通用，与图5 F5POOL 同口径）：近8周 w8 全为0的型号不入池 */
function zeroSales9(m) { const w=(m.gro||{}).w8||[]; return w.length && w.every(v=>!v); }

/* ============ 新品（近4周滚动窗口首次出现） ============ */
/* 当前近4周窗口：META.week_end 为末周(n)，起点 = n-3。例：期末37 → 新品=首现于26W34~37 */
const _newEnd = D.meta && D.meta.week_end || 37;
const _newEndNum = 202600 + (+_newEnd);
const _newStartNum = _newEndNum - 3;
/* 真新品：在近4周滚动窗口内首次出现，且此前从未出现在AVC（first_num 落在窗口内即满足） */
function isNew(m){
  const fn = (m.gro||{}).first_num;
  return fn!=null && fn>=_newStartNum && fn<=_newEndNum;
}
/* 新品近4周累计销量（w8 为 26W30-37，近4周 = 后4个元素） */
function new4w(m){ const w=(m.gro||{}).w8||[]; return w.slice(-4).reduce((a,b)=>a+(b||0),0); }
/* 新品准入门槛（图1-4、图5共用）：常规 近4周累计>50；98吋+ 不再豁免，需 近4周累计>10 */
function newGate(m){ const s=m.size||0; if(s>=93) return new4w(m)>10; return new4w(m)>50; }
/* 新品近4周是否在售（至少1周有销量）——用于确认新品不被"近8周全0"误删，替代 zeroSales9 */
function newActive(m){ return new4w(m)>0; }
/* 是否按"新品口径"放入该图（图1-4 通用）：新品 hit 后不再要求旧门槛；旧门槛仍由 pass/meta gate 控制 */
function newEligible(m){ return isNew(m); }
/* 新品保护期（读自 Python 侧持久化的曾达标标记）：= 曾达标新品且首现后20周内 */
function f5Guard(m){ return !!m.nguard; }
/* 绿点：保护期内、且已摘「新品」标（到期16周拓展段）才显示绿点；新品期仍显示绿色 chip */
function f5Dot(m){ return f5Guard(m) && !isNew(m); }
/* 绿点 HTML：保护期内且已摘「新品」标时，在型号右上角打绿色圆点；新品期仍显示绿色「新品」chip */
function gDot(m){ return f5Dot(m) ? '<i class="gdot" title="新品保护期：豁免清库与近8周全0，首现后20周内"></i>' : ''; }
/* 清库 / 近8周全0 豁免：新品 + 保护期全部豁免 */
function protEx(m){ return isNew(m) || f5Guard(m); }

/* 去掉型号开头的尺寸前缀：75A4H→A4H；SKS30→SKS30（无尺寸前缀保留） */
const _SZ_PREFIX = new Set(['32','40','43','50','55','60','65','70','75','77','85','88','98','100','110','120','163']);
function stripSizeModel(model){
  const s = String(model||'');
  for (let n = 3; n >= 2; n--) {           // 先试3位（100/110/163）再2位
    const pre = s.slice(0, n);
    const rest = s.slice(n);
    if (_SZ_PREFIX.has(pre) && rest && /^[A-Za-z]/.test(rest)) return rest;
  }
  return s;
}
/* 真激光判定：仅按屏幕技术 tech 含 LASER（激光投影）确定为激光电视。
   AVC 的 series='激光电视' 实为 100吋+ 大屏产品线的命名（如 100A9H/163QX 等 MiniLED/LCD 液晶），不算激光，
   不作为剔除依据。新品板块因此不会引入真激光，同时保留液晶大屏新品。 */
function isNewLaser(m){
  const t = m.tech||'';
  return t.indexOf('LASER')>=0;
}
/* 新品板块显示名（分组键）：4品牌用产品名，非4品牌用去尺寸型号名 */
function newGroupOf(m){
  const isProduct = BRANDS_USE_PRODUCT.has(m.brand);
  if (isProduct) return (m.product_name || stripSizeModel(m.model));
  return stripSizeModel(m.model);
}

/* 新品上市板块：列出全部真新品（含未达标/待观察），
   同品牌同「分组键」合并为一张系列卡片（如 小米 L75MD-SP/L85MD-SP → S Pro 27款）。
   分组键：4品牌(雷鸟/Vidda/小米/红米)用 product_name（缺省回退去尺寸型号名），其余品牌用「不带尺寸的型号名」。
   hover 悬停显示该系列下全部 AVC 原始型号。同系列内平铺各尺寸卡片。 */
function fillNewSec() {
  const sec = document.getElementById('newsec');
  if (!sec) return;
  const all = M.filter(m => isNew(m) && !isNewLaser(m) && newGate(m));   // 仅展示达标新品；激光电视新品永不引入
  if (!all.length) { sec.style.display='none'; return; }
  sec.style.display='';
  const byBrand = {};
  all.forEach(m => byBrand[m.brand]=(byBrand[m.brand]||0)+1);
  const brandDist = Object.keys(byBrand).map(b=>`${b} ${byBrand[b]}`).join(' / ');
  document.getElementById('newcnt').textContent =
    `共 ${all.length} 款 / ${new Set(all.map(m=>newGroupOf(m)+'|'+m.brand)).size} 系列 · 均为达标新品　品牌分布：${brandDist}`;
  document.getElementById('newsub').textContent =
    `新品定义：近4周滚动窗口(26W${_newEndNum-202600-3}–W${_newEndNum-202600})内首次在 AVC 出现（此前从未出现），仅展示达标新品。同品牌同系列多尺寸合并展示，悬停看 AVC 原始型号。达标门槛：常规 近4周累计>50台；98吋+ 近4周累计>10台。下期数据滚动后，超出近4周窗口的新品将摘除「新品」标，转为常规型号继续展示。`;
  const grid = document.getElementById('newgrid');
  // 按 (品牌,分组键) 聚合；组内保持按达标优先、近4周销量降序
  const grp = {};
  all.forEach(m => {
    const k = m.brand + '\u0001' + newGroupOf(m);
    (grp[k] = grp[k] || []).push(m);
  });
  grid.innerHTML = Object.keys(grp)
    .sort((a,b) => (grp[b].some(newGate)-grp[a].some(newGate)) || (grp[b].reduce((s,m)=>s+new4w(m),0) - grp[a].reduce((s,m)=>s+new4w(m),0)))
    .map(k => {
      const ms = grp[k];
      const m0 = ms[0];
      const nm = newGroupOf(m0);                 // 分组显示名：产品名/去尺寸型号名
      // hover：列出该系列全部 AVC 原始型号
      const modelList = ms.map(m => `${m.model}（${m.size}吋，近4周 ${new4w(m)} 台）`).join('\n');
      const title = `${m0.brand} ${nm}\n型号：${ms.map(m=>m.model).join(' / ')}\n${modelList}`;
      // 组内尺寸小卡
      const items = ms
        .slice()
        .sort((a,b) => (newGate(b)-newGate(a)) || (new4w(b)-new4w(a)))
        .map(m => {
          const w4 = new4w(m), sz = m.size||0;
          const modelTip = `型号 ${m.model} · ${m.size}吋 · AVC 均价 ¥${m.avg_price||'—'}　首现 ${m.gro.first_week||'?'}`;
          const exNote = sz>=93?'<span class="ex">(98吋+ 达标&gt;10台)</span>':'';
          return `<div class="nitem">
            <div class="itop"><span class="sz">${m.size}吋</span><span class="mrl" title="${modelTip}">${m.model}</span>
              <span class="tag-done">达标</span></div>
            <div class="imet">首现 ${m.gro.first_week||'?'} · 近4周 <b>${w4}</b> 台 · 均价 <b>¥${m.avg_price||'—'}</b> ${exNote}</div>
          </div>`;
        }).join('');
      return `<div class="newcard">
        <div class="top"><span class="bn" title="${title}">${m0.brand} ${nm}${ms.length>1?` <em class="ncnt">×${ms.length}尺寸</em>`:''}</span>
          <span class="tag-done">达标</span>
          <span class="ngchip">新品</span></div>
        <div class="nlist">${items}</div>
      </div>`;
    }).join('');
}
/* 图1-4 通用清库判定（按尺寸差异化，与图5 f5Clear 兼容）：
   大尺寸(>50吋)：合计<100 或 连续≥4周<10，高价位例外阈值10
   小尺寸(≤50吋)：合计<500 或 连续≥4周<50；豁免 32吋 均价>1500、40/43吋 均价>2000、50吋 均价>2500（豁免需近8周有动销） */
function isClear9(m){
  const w=(m.gro||{}).w8||[]; if(!w.length) return null;
  if(protEx(m)) return null;   // 新品+保护期豁免清库打标
  const sz = m.size||0;
  const sum = w.reduce((a,b)=>a+(b||0),0);
  let T, ex=false;
  if (sz>50) {
    // 大尺寸高价位例外档（内联，避免依赖 F5SZ_OF 的 TDZ）：与图5 f5Excp 同口径
    const z = sz<=60?'55':sz<=69?'65':sz<=79?'75':sz<=92?'85':'98+';
    const p = m.avg_price||0;
    if (z==='98+' && p < 16000) {
      // 98吋+ 低价(<16000)：累计<300 且 连续≥4周单周<10 双条件同时命中才判清库
      let run=0, mxr=0;
      for(const v of w){ run = (v<10) ? run+1 : 0; if(run>mxr) mxr=run; }
      if (!(sum < 300 && mxr >= 4)) return null;
      return { reason:`近8周累计${sum}台 <300台 且连续${mxr}周<10台/周`, ex:false, T:300 };
    }
    if (z==='98+') { ex = true; T = 10; }
    else { ex = z==='55'?p>=5000: z==='65'?p>=6000: z==='75'?p>=7000: p>=9000; T = ex?10:100; }
  } else {
    // 小尺寸豁免：32吋(≤35) 均价>1500；40/43吋(36-46) 均价>2000；50吋(47-50) 均价>2500。豁免前提是近8周有动销(sum>0)
    const p = m.avg_price||0;
    const exempt = sz<=35 ? p>=1500 : sz<=46 ? p>=2000 : p>=2500;
    if (exempt && sum>0) return null;
    T = 500;
  }
  let run=0, maxrun=0;
  for(const v of w){ run = (v<T) ? run+1 : 0; if(run>maxrun) maxrun=run; }
  const hitA = maxrun>=4, hitB = sum<T;
  if(!hitA && !hitB) return null;
  const rs=[];
  if(hitB) rs.push(`近8周合计${sum}台 <${T}台`);
  if(hitA) rs.push(`连续${maxrun}周 <${T}台/周`);
  return { reason: rs.join('；'), ex, T };
}
/* 单一图：全局筛选 + 该图独立品牌集合 */
function figList(fig) {
  const bs = S.b[fig];
  return M.filter(m => bs.has(m.brand) && pass(m) && (!zeroSales9(m) || protEx(m)) && isClearShow(fig,m));
}

/* 「移除清库机」过滤器：fig 图开启(S.clr[fig]===1)时，清库型号(m)被排除；未开启则全保留 */
function isClearShow(fig, m) {
  if (!S.clr[fig]) return true;
  return !isClear9(m);
}

/* 清库判定文字（按钮旁定义说明）。两口径共用，标注尺寸差异。 */
const CLR_DEF_TEXT = '清库型号定义：<b>大尺寸</b>(>50吋) 近8周合计销量&lt;100台 或连续≥4周&lt;10台（高端价格带豁免阈值10台）；<b>98吋+</b> 均价≥16000元阈值10台、均价&lt;16000元需累计&lt;300台且连续≥4周单周&lt;10台；<b>小尺寸</b>(≤50吋) 近8周合计&lt;500台 或连续≥4周&lt;50台。小尺寸豁免：32吋 均价&gt;1500元、40/43吋 均价&gt;2000元、50吋 均价&gt;2500元（豁免型号需近8周有动销）。近8周0销量型号始终不收。';

/* 各图「移除清库机」开关条：注入到指定宿主元素之后；开启时图内清库型号不显示 */
function buildClearCtl(fig, hostId) {
  const host = document.getElementById(hostId);
  if (!host) return;
  let box = document.getElementById('clrctl'+fig);
  if (!box) {
    box = document.createElement('div');
    box.className = 'clrctl';
    box.id = 'clrctl'+fig;
    host.insertAdjacentElement('afterend', box);
    box.innerHTML = '<span class="tag ctl-label">清库</span>'
      + '<button class="clrbtn" data-cf="'+fig+'">移除清库机</button>'
      + '<span class="clr-def">'+CLR_DEF_TEXT+'</span>';
  }
  box.querySelector('.clrbtn').classList.toggle('on', !!S.clr[fig]);
  box.querySelector('.clrbtn').textContent = S.clr[fig] ? '已移除清库机（点按恢复）' : '移除清库机';
  const btn = box.querySelector('.clrbtn');
  if (!btn._bound) {
    btn._bound = true;
    btn.onclick = () => {
      S.clr[fig] = S.clr[fig] ? 0 : 1;
      /* 依图类型重渲染 */
      if (fig === 5) { f5BrandFilter(); try { f5Grid(); f5Draw(); } catch(e){} }
      else { refreshFig(fig); }
      buildClearCtl(fig, hostId);
    };
  }
}

function seriesRows(list) {
  const g = new Map();
  for (const m of list) {
    // 新品豁免原门槛：只要满足新品门槛(newGate)和尺寸有归属列(col!=null)就可以入池；保护期(摘标后)同样豁免
    if (!m.fig1_ok && !(isNew(m) && newGate(m) && m.fig1_col!=null) && !f5Guard(m)) continue;
    const k = m.brand + '||' + m.series;
    if (!g.has(k)) g.set(k, {brand:m.brand, series:m.series, vol:0, n:0, cells:new Map()});
    const r = g.get(k);
    r.vol += m.cum_vol; r.n++;
    // 图1 列 = 参考报告的 10 个尺寸列（60吋已并入 55）
    const col = m.fig1_col;
    if (col == null) continue;
    if (!r.cells.has(col)) r.cells.set(col, []);
    r.cells.get(col).push(m);
  }
  const rows = [...g.values()];
  // 计算每个系列的综合加权均价 wavg，以及「主力尺寸(销量最大列)同列均价 colAvg」
  // —— 海信/TCL/创维以 colAvg 作为高低端排序依据，避免"只有超大尺寸故绝对价高"的偏差
  for (const r of rows) {
    let wv = 0, ap = 0, bestCol = null, bestVol = -1;
    for (const [col, arr] of r.cells) {
      let cv = 0, ca = 0;
      for (const m of arr) { cv += m.cum_vol; ca += m.avg_price * m.cum_vol; }
      wv += cv; ap += ca;
      if (cv > bestVol) { bestVol = cv; bestCol = col; }
    }
    r.wavg = wv ? ap / wv : 0;
    r.mainSize = bestCol; r.colAvg = 0;
    const tar = bestCol != null ? r.cells.get(bestCol) : null;
    if (tar) { let tv = 0, ta = 0; for (const m of tar) { tv += m.cum_vol; ta += m.avg_price * m.cum_vol; } r.colAvg = tv ? ta / tv : 0; }
  }
  // 排序：先按品牌（8 品牌按既定顺序）；酷开品牌内 P 系（P开头系列）按第二个字母降序置顶，
  // 非P系排在 P 系之后按 wavg 降序；海信/TCL/创维 按用户指定的固定系列顺序（RANK）；
  // 其余品牌按综合加权均价 wavg 降序
  const RANK = {
    海信: {'E8 旗舰系':1,'E7 高端系':2,'D7 系':3,'E5 中端系':4,'D6/D5 系':5,'E52 系':6,'D3 系':7,'E2/E3 入门系':8},
    TCL: {'Q10 系':1,'Q9 系':2,'Q7 系':3,'Q6 系':4,'T7 系':5,'壁纸电视':6,'T6 系':7,'T5 系':8,'V8 系':9,'其他':10,'LIFE VISION':11},
    创维: {'A9/A10 高端系':1,'A8 系':2,'A7 系':3,'A6 系':4,'A28 系':5,'A5 系':6,'A4 系':7,'A3 系':8,'A 系入门':9,'其他':10},
  };
  const pTier = s => /^P\d/.test(s||'') ? +s[1] : -1;   // P系列第二个字母档位；非P系=-1
  rows.sort((a,b) => {
    const bi = D.brands.indexOf(a.brand)-D.brands.indexOf(b.brand);
    if (bi) return bi;
    if (a.brand==='酷开' && b.brand==='酷开') {
      const ta = pTier(a.series), tb = pTier(b.series);
      if (ta !== tb) return (tb===-1?1:tb) - (ta===-1?-1:ta);   // P档位降序，档位-1(非P)放最后
      return b.wavg - a.wavg;
    }
    if (RANK[a.brand]) {
      const ra = RANK[a.brand][a.series] ?? 999, rb = RANK[a.brand][b.series] ?? 999;   // 手工固定顺序
      return ra - rb || b.wavg - a.wavg;
    }
    return b.wavg - a.wavg;
  });
  if (S.mainOnly) return rows.slice(0, 40);
  return rows;
}

function heatColor(brand, vol, max) {
  if (!vol || !max) return '#fafbfc';
  const t = Math.pow(vol/max, 0.42);
  const c = CO[brand] || '#0E7CE8';
  const r = parseInt(c.slice(1,3),16), g = parseInt(c.slice(3,5),16), b = parseInt(c.slice(5,7),16);
  const mix = v => Math.round(250 + (v-250)*t*0.82);
  return `rgb(${mix(r)},${mix(g)},${mix(b)})`;
}
function txtOn(t) { return t > 0.52 ? '#fff' : '#1f2328'; }

/* 原生刷新率数值：去掉 "4K "、"(300Hz)" 等分辨率前缀与电竞插值后缀，只留原生 Hz 数 */
function nativeHz(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return String(v);
  const m = String(v).match(/(\d+(?:\.\d+)?)\s*Hz/);
  return m ? m[1] : '';
}

/* 悬停参数文本：只显示值、不显示字段名；平凡值（普通液晶/LED、无分区、无抗反射）不显示；
   刷新率带 Hz 且前带分辨率（如 "4K 120Hz"）；音响带声道（如 "2.0声道"） */
function hoverParams(m) {
  const segs = [];
  if (m.size > 50) {
    // >50吋：进阶技术 / 分区档 / 原生刷新率 / 内存 / 抗反射 / 音响（六维，无分辨率）
    const t = m.tech || '';
    if (strongTech(t)) segs.push(t);
    const pExact = m.part;   // 精确分区数（如 720、4048），无则空
    const hasExact = pExact != null && pExact !== '' && pExact !== '无' && Number(pExact) > 0;
    const pb = m.part_band || '';
    if (hasExact) {
      segs.push(pExact + '分区');            // 有精确值 → 显示精确分区数
    } else if (pb && pb !== '无分区') {
      segs.push(partLabel(pb));              // 无精确值才回退档位
    }
    const hz = nativeHz(m.refresh_hz);
    if (hz) segs.push(hz + 'Hz');
    if (m.mem) segs.push(m.mem);
    if (m.anti && m.anti !== '无') segs.push(m.anti);
    if (m.audio && m.audio !== '(免)') segs.push(m.audio);
  } else {
    // ≤50吋：只显示 分辨率 / 内存 / 能效（不显示刷新率/分区/抗反射/音响）
    const res = m.res_band || (m.res || '').replace(/[\d.]+\s*Hz.*$/i, '').trim();
    if (res) segs.push(res);
    if (m.mem) segs.push(m.mem);
    if (m.energy && m.energy !== '' && m.energy !== '待补' && m.energy !== '不确定') segs.push(m.energy);
  }
  return segs.join('｜');
}

/* 屏幕技术判强：忽略大小写与空格/连字符差异（'Mini LED'↔'MiniLED'、'QD-Mini LED'↔'QD-MiniLED'） */
function strongTech(t) {
  if (!t || !D.techStrong) return false;
  const n = String(t).toLowerCase().replace(/[\s-]/g, '');
  return D.techStrong.some(s => String(s).toLowerCase().replace(/[\s-]/g, '') === n);
}
/* 图5「技术」列：普通液晶类（LCD/LED/普通液晶/直下式）统一显示「LED」且不高亮；
   进阶技术（MiniLED 系/QLED/OLED/量子点）写紧凑原名并高亮；空值标「待补」 */
/* 屏幕技术统一名：OLED / RGB-Mini LED / SQD-Mini LED / BGB-Mini LED / QD-Mini LED / Mini LED / QLED / LED */
function f5TechNorm(t) {
  const s = String(t || '').trim();
  if (!s) return '';
  if (s.includes('OLED')) return 'OLED';
  if (s.includes('RGB-Mini') || s.includes('RGB Mini')) return 'RGB-Mini LED';
  if (s.includes('SQD')) return 'SQD-Mini LED';
  if (s.includes('BGB-Mini') || s.includes('BGB Mini')) return 'BGB-Mini LED';
  if (s.includes('QD-Mini')) return 'QD-Mini LED';
  if (s.includes('Mini')) return 'Mini LED';
  if (s.includes('QLED') || s.includes('量子点')) return 'QLED';
  return 'LED'; // LCD / LED / 普通液晶 / 直下式
}
function f5Tech(m) {
  const n = f5TechNorm(m && m.tech ? String(m.tech).trim() : '');
  if (!n) return '<span class="na">待补</span>';
  if (n === 'LED') return 'LED';
  return `<b class="up" title="屏幕技术：${n}">${n}</b>`;
}
/* 原生刷新率：统一加 Hz 单位 */
function f5Hz(v) {
  if (v == null || String(v).trim()==='') return '<span class="na">待补</span>';
  const s = String(v);
  return /Hz$/i.test(s) ? s : s + 'Hz';
}
/* 内存格式：3.0/4.0 → 3/4（整数 RAM），1.5 保留 */
function memFmt(mem) {
  if (!mem) return '<span class="na">待补</span>';
  return String(mem).replace(/\b(\d+)\.0(?=[+×Xx]|$)/g, '$1');
}
/* 内存档位：字段缺失一律按「待补」处理（否则落回白底，造成「残缺」观感） */
function memBandOf(m) { return m.mem_band || '待补'; }

/* 卡片 chip 渲染：
   - 主名：4 品牌用 product_name（雷鸟/Vidda/小米/红米），其余用 model
   - tech 徽章：仅进阶技术（Mini LED 系/OLED/QLED/量子点/激光）显示；
     普通液晶/LED 及无技术数据一律不标（用户要求：非 Mini LED 的普通 LED 不标注）
   - 产品名卡片不显示型号副标（简化信息）
   - 内存颜色：单个卡片淡色底；待补(无 mem_band) 用虚线边框 + 待补色，避免「残缺」观感
   - hover 显示参考报告 chip 完整参数文本（>50吋/≤50吋 各自有专属字段） */
function chip(m, showAnti) {
  const c = CO[m.brand] || '#888';
  const isProduct = BRANDS_USE_PRODUCT.has(m.brand) && m.product_name;
  const main = isProduct ? m.product_name : m.model;
  const anti = m.anti && m.anti !== '无' ? `<span class="badge">${m.anti}</span>` : '';
  const sv = showAnti ? anti : '';
  // tech 徽章：仅进阶技术显示；其余（普通 LED/液晶/无数据）不标注
  let techBadge = '';
  if (strongTech(m.tech)) {
    const t = m.tech;
    const short = (D.techLabel && D.techLabel[t]) || t;
    techBadge = `<span class="badge tech hi" title="屏幕技术：${t}">${short}</span>`;
  }
  // 清库徽章（图1-4 通用打标，图5 独立处理）
  const cl9 = isClear9(m);
  const cl9tip = cl9 ? '清库:' + cl9.reason + (cl9.ex ? '（高价位例外，阈值' + cl9.T + '台）' : '') : '';
  const clearBadge = cl9 ? '<span class="badge clearb" title="' + cl9tip.replace(/"/g,'&quot;') + '">清库</span>' : '';
  // 新品徽章（绿色；新品豁免清库故二者互斥）
  const newBadge = isNew(m) ? '<span class="ngchip">新品</span>' : '';
  // 内存色 + 待补虚线边框
  const memB = memBandOf(m);
  const mbg = (D.memCardBg && D.memCardBg[memB]) || '#fff';
  const memNa = memB === '待补' ? ' memna' : '';
  // hover 文本：首行 = 品牌/型号/尺寸 + AVC 累计销量/均价；第二段 = 仅参数值（无字段名）
  const hp = hoverParams(m);
  const hover = `${m.brand} ${m.model} ${m.size}吋｜${isProduct?'产品名：'+m.product_name+' ｜':''}累计 ${fmt(m.cum_vol)} 台｜均价 ¥${m.avg_price}` +
    (hp ? '\n\n' + hp : '');
  return `<div class="mc${memNa} ${S.sel.includes(m.id)?'sel':''}" data-id="${m.id}" title="${hover.replace(/"/g,'&quot;')}">
    <div class="mbar" style="background:${c}"></div>
    <div class="mbody" style="background:${mbg}">
      <div class="mtop"><span class="md" style="background:${c}"></span><span class="mnm">${main}</span>${gDot(m)}${techBadge}${clearBadge}${newBadge}</div>
      <div class="mbot"><span>${m.size}吋</span><span>${fmt(m.cum_vol)}台</span>${sv}</div>
    </div></div>`;
}

/* 四图共用同一最小宽度基准，保证任何视口下外框宽度完全一致 */
const LABW = 168, MCELL = 130;
function ftHead(label, cols) {
  S.ncols.push(cols.length);
  let h = `<table class="ft" style="--nc:${cols.length}">`;
  h += '<colgroup><col class="lab">' + '<col class="d">'.repeat(cols.length) + '</colgroup>';
  h += '<thead>';
  h += `<tr><th style="text-align:left">${label}</th>`;
  for (const c of cols) h += `<th>${c}</th>`;
  h += '</tr></thead><tbody>';
  return h;
}

/* 四图渲染完后统一取最大列数，写入 --mxn，使四表 min-width 相同 */
function syncTableWidth() {
  document.documentElement.style.setProperty('--mxn', Math.max(1, ...S.ncols));
  S.ncols = [];
}

/* ===== 四图表头浮动吸顶（仿参考报告：克隆 thead 到 fixed 层，不压缩表格高度） =====
   四图表格由 JS 动态渲染，故每次 renderAll 后需重建。                        */
let FHEADS = [];
function buildFloatHeads() {
  document.querySelectorAll('.float-head').forEach(e => e.remove());
  FHEADS = [];
  document.querySelectorAll('table.ft, table.f5tbl').forEach(tb => {
    if (!tb.tHead) return;
    const wrap = tb.parentElement;
    if (!wrap) return;
    const box = document.createElement('div');
    box.className = 'float-head';
    if (wrap.id) box.dataset.fid = wrap.id;   // 便于按图定位
    const clone = tb.cloneNode(false);
    clone.appendChild(tb.tHead.cloneNode(true));
    box.appendChild(clone);
    document.body.appendChild(box);
    FHEADS.push({tb:tb, wrap:wrap, box:box, clone:clone, dirty:1});
    // wrap 是持久节点，避免重复绑定
    if (!wrap.__fhBound) { wrap.addEventListener('scroll', updateFloatHead); wrap.__fhBound = 1; }
  });
  updateFloatHead();
}
function syncFHWidth(f) {
  const src = f.tb.tHead.rows[0].cells, dst = f.clone.tHead.rows[0].cells;
  for (let i = 0; i < src.length && i < dst.length; i++) {
    const w = src[i].getBoundingClientRect().width;
    dst[i].style.width = w + 'px';
    dst[i].style.minWidth = w + 'px';
    dst[i].style.maxWidth = w + 'px';
  }
  f.clone.style.width = f.tb.getBoundingClientRect().width + 'px';
}
function updateFloatHead() {
  // 源表若已被重新渲染替换（innerHTML 覆盖），重建浮动层，避免引用失效节点
  if (FHEADS.length && !FHEADS[0].tb.isConnected) { buildFloatHeads(); return; }
  for (const f of FHEADS) {
    const r = f.tb.getBoundingClientRect();
    const wr = f.wrap.getBoundingClientRect();
    if (r.height === 0 || wr.top > 0 || wr.bottom < 40) { f.box.style.display = 'none'; f.dirty = 1; continue; }
    if (f.box.style.display !== 'block' || f.dirty) { f.box.style.display = 'block'; syncFHWidth(f); f.dirty = 0; }
    f.box.style.left = wr.left + 'px';
    f.box.style.width = wr.width + 'px';
    f.clone.style.transform = 'translateX(' + (-f.wrap.scrollLeft) + 'px)';
  }
}
window.addEventListener('scroll', updateFloatHead, true);
window.addEventListener('resize', function(){
  FHEADS.forEach(f => { f.box.style.display = 'none'; f.dirty = 1; });
  updateFloatHead();
});

/* 图1 单元格：完全采用参考报告「海信TCL产品布局分析」的明细规则
   - 行头：品牌色 4px 左边条 + 系列名 + 型号数
   - 表头：每列标注该尺寸的型号数
   - 单元格：m-count 型号数 + 逐款 m-row（去尺寸前缀名 + 销量），按销量降序
   - 空单元格用 .empty 浅灰底                                                */
const F1_MAXROWS = 8;
function shortName(m) {
  // 雷鸟 / Vidda / 小米 / 红米：直接用对照表的产品名称（含「-渠道」后缀，按用户 09-04 要求原样保留）
  if (BRANDS_USE_PRODUCT.has(m.brand) && m.product_name) {
    return String(m.product_name);
  }
  return String(m.model).replace(/^\d{2,3}/, '') || m.model;
}
function renderFig1(list) {
  const rows = seriesRows(list);
  const max = Math.max(1, ...rows.map(r => Math.max(0, ...[...r.cells.values()].map(a => a.reduce((s,x)=>s+x.cum_vol,0)))));
  const cols = D.fig1Cols;
  const colCnt = new Map();
  for (const r of rows) for (const [c, arr] of r.cells) colCnt.set(c, (colCnt.get(c)||0) + arr.length);
  const head = cols.map(c => `${D.fig1Lab[String(c)]}<br><span class="cnt-label">${colCnt.get(c)||0}个</span>`);
  let h = ftHead('系列<br><span class="cnt-label">(型号数)</span>', head);
  for (const r of rows) {
    const c = CO[r.brand] || '#0E7CE8';
    h += `<tr><td class="rowhead"><div class="rl" style="border-left-color:${c}"><span class="bn">${r.brand}</span><span class="sn">${r.series}</span><span class="cnt">${r.n}款</span></div></td>`;
    for (const col of cols) {
      const arr = r.cells.get(col);
      if (!arr || !arr.length) { h += '<td><div class="cell empty"></div></td>'; continue; }
      const v = arr.reduce((a,x)=>a+x.cum_vol,0);
      const t = Math.pow(v/max, 0.42);
      const bg = heatColor(r.brand, v, max);
      const fg = txtOn(t);
      arr.sort((a,b)=>b.cum_vol-a.cum_vol);
      const shown = arr.slice(0, F1_MAXROWS);
      const more = arr.length > F1_MAXROWS ? `<div class="m-more">还有 ${arr.length-F1_MAXROWS} 款</div>` : '';
      h += `<td><div class="cell f1" style="background:${bg};color:${fg}" data-brand="${r.brand}" data-series="${r.series}" data-col="${col}">
        <div class="m-list"><span class="m-count">${arr.length}个</span>
        ${shown.map(m=>{
          const cl1 = isClear9(m);
          const clT = cl1 ? '清库:'+cl1.reason+(cl1.ex?'（高价位例外，阈值'+cl1.T+'台）':'') : '';
          const clB = cl1 ? `<i class="m-clr" title="${clT.replace(/"/g,'&quot;')}">清库</i>` : '';
          const ngB = isNew(m) ? '<i class="m-ng">新</i>' : '';
          return `<div class="m-row" title="${m.brand} ${m.model} ${m.size}吋｜累计 ${fmt(m.cum_vol)} 台｜均价 ¥${m.avg_price.toLocaleString()}${hoverParams(m)?'\n\n'+hoverParams(m):''}" data-id="${m.id}"><span class="m-name">${shortName(m)}</span>${gDot(m)}${ngB}${clB}<span class="m-vol">${fmt(m.cum_vol)}</span></div>`;
        }).join('')}
        ${more}</div></div></td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  document.getElementById('fig1body').innerHTML = h;
  document.getElementById('fig1cnt').textContent = `${rows.length} 个系列 · ${list.filter(m=>m.fig1_ok).length} 款 · 单格销量峰值 ${fmt(max)} 台`;
  const k1 = [];
  for (const r of rows) for (const col of r.cells.keys()) k1.push(r.brand + '|' + r.series + '|' + col);
  updateExpBtn(1, k1);
}

/* 图2 列：98-100吋 与 100吋以上 合并为「98吋+」（用户 2026-09-04） */
function fig2ColOf(m) {
  return (m.size_bucket === '98-100吋' || m.size_bucket === '100吋以上') ? '98吋+' : m.size_bucket;
}
const FIG2_COLS = ['50吋及以下', '55吋', '65吋', '75吋', '85吋', '98吋+'];

function renderFig2(list) {
  const g = new Map();
  for (const m of list) {
    if (!m.fig1_ok && !(isNew(m) && newGate(m) && m.fig1_col!=null) && !f5Guard(m)) continue;   // 图2 跟随图1，新品与保护期豁免门槛
    const b = beltOf(m), s = fig2ColOf(m);
    if (!g.has(b)) g.set(b, new Map());
    if (!g.get(b).has(s)) g.get(b).set(s, {n:0,v:0,arr:[]});
    const c = g.get(b).get(s);
    c.n++; c.v += m.cum_vol; c.arr.push(m);
  }
  const belts = D.beltOrder.concat(['待补']).filter(b => g.has(b));
  const max = Math.max(1, ...[...g.values()].flatMap(r => [...r.values()].map(c => c.v)));
  let h = ftHead('均价带', FIG2_COLS);
  for (const b of belts) {
    h += `<tr><td><div class="rl"><span class="sn">${b}</span></div></td>`;
    for (const s of FIG2_COLS) {
      const c = g.get(b).get(s);
      if (!c) { h += '<td><div class="cell empty"></div></td>'; continue; }
      const t = Math.pow(c.v/max, 0.42);
      const mix = v => Math.round(250 + (v-250)*t*0.8);
      const bg = `rgb(${mix(110)},${mix(110)},${mix(110)})`;
      // 方案 A：品牌份额堆叠条（TOP3 + 其他），条色 = 品牌色
      const byB = new Map();
      for (const m of c.arr) byB.set(m.brand, (byB.get(m.brand)||0) + m.cum_vol);
      const bl = [...byB.entries()].sort((a,b2)=>b2[1]-a[1]);
      const top = bl.slice(0,3);
      const othV = bl.slice(3).reduce((s2,x)=>s2+x[1],0);
      const segs = othV>0 ? top.concat([['其他',othV]]) : top;
      const bar = segs.map(([bn,v]) => `<i style="width:${(v/c.v*100).toFixed(1)}%;background:${bn==='其他'?'#b9c0ca':(CO[bn]||'#b9c0ca')}"></i>`).join('');
      const hover = '品牌构成（' + fmt(c.v) + '台）：\n' + bl.map(([bn,v]) => `${bn} ${fmt(v)}台 · ${(v/c.v*100).toFixed(0)}%`).join('\n');
      h += `<td><div class="cell hasbar" style="background:${bg};color:${t>0.52?'#fff':'#1f2328'}" data-belt="${b}" data-size="${s}" title="${hover.replace(/"/g,'&quot;')}">
        <span class="vn"><span class="v">${fmt(c.v)}</span><span class="n">${c.n}款</span></span><span class="sbar">${bar}</span></div></td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  document.getElementById('fig2body').innerHTML = h;
  let tot = 0; for (const r of g.values()) for (const c of r.values()) tot += c.n;
  document.getElementById('fig2cnt').textContent = `${belts.length} 个价格带 · ${tot} 款`;
  const k2 = [];
  for (const b of belts) for (const s of FIG2_COLS) if (g.get(b) && g.get(b).get(s)) k2.push('B|' + b + '|' + s);
  updateExpBtn(2, k2);
}

function bandVal(m, key) { return m[key] || '待补'; }
/* 参数门槛（近20周 入选口径）：图3 主流>3000 / 98吋+>500；图4 ≤50吋 近20周>1000 */
function gateVol(m) { const v20 = m.vol20||0; return m.size < 98 ? v20 > 3000 : v20 > 500; }
const hzKey = m => m.refresh_hz ? m.refresh_hz + 'Hz' : '待补';

function toggleF(set, key) {
  if (!set) return new Set([key]);  // null = 全选 → 显式选择此 key
  if (set.has(key)) { set.delete(key); if (!set.size) return null; }
  else set.add(key);
  return set;
}
function visibleF(set, key) {
  if (!set) return true;            // null = 全选
  return set.has(key);
}
/* 图3 刷新率列分组：120/132 并作「120-132Hz」，165/170 并作「165-170Hz」，其余原值 */
function hzBandOf(h) {
  if (h == null) return '待补';
  if (h === 120 || h === 132) return '120-132Hz';
  if (h === 165 || h === 170) return '165-170Hz';
  return h + 'Hz';
}
/* 单图内筛选后局部重绘：
   - 只重渲染该图，其它三图不受影响
   - S.ncols 清零：--mxn 保持 renderAll 时算好的值，避免四图宽度被单图列数带偏
   - 重绘后源表被 innerHTML 覆盖，浮动表头必须重建                     */
function refreshFig(fig) {
  S.ncols = [];
  const list = figList(fig);
  if (fig === 1) renderFig1(list);
  else if (fig === 2) renderFig2(list);
  else if (fig === 3) renderFig3(list);
  else renderFig4(list);
  S.ncols = [];
  bindCells();
  buildFloatHeads();
}

function colSortHZ(a,b) {
  const na = a==='待补' ? Infinity : (parseInt(a)||0), nb = b==='待补' ? Infinity : (parseInt(b)||0);
  return na - nb;
}
/* 二轴网格单元格：同品牌型号折叠成一组；组头=品牌名+代表型号(销量第一)+数量 */
function cellGroups(arr, showAnti, keyPrefix, set) {
  const bvol = new Map();
  for (const m of arr) bvol.set(m.brand, (bvol.get(m.brand)||0) + m.cum_vol);
  const brands = [...bvol.keys()].sort((a,b)=>bvol.get(b)-bvol.get(a));
  let h = '';
  for (const bn of brands) {
    const garr = arr.filter(m=>m.brand===bn).sort((a,b)=>b.cum_vol-a.cum_vol);
    const rep = garr[0];
    const k = keyPrefix + '||' + bn;
    const open = set.has(k);
    const tip = (shortName(rep)+'　'+hoverParams(rep)).replace(/"/g,'&quot;');
    h += `<div class="bgg"><div class="bghead" data-gk="${k}" style="border-left-color:${CO[bn]||'#888'}" title="点击展开/收起 ${bn} 的 ${garr.length} 款">
        <span class="gtri${open?' on':''}">${open?'▾':'▸'}</span><span class="gbn">${bn}</span>
        <span class="grep" title="${tip}">${shortName(rep)}</span><span class="gcnt">${garr.length}</span></div>
      <div class="bgbody"${open?'':' style="display:none"'}><div class="gridbox-in">${garr.map(m=>chip(m,showAnti)).join('')||''}</div></div></div>`;
  }
  return h;
}
function bindGroupHeads(fig) {
  const body = document.getElementById(fig===3 ? 'fig3body' : 'fig4body');
  if (!body) return;
  body.querySelectorAll('[data-gk]').forEach(el => el.onclick = e => {
    e.stopPropagation();
    const set = fig===3 ? S.f3g : S.f4g, k = el.dataset.gk;
    if (set.has(k)) set.delete(k); else set.add(k);
    refreshFig(fig);
  });
}
function toggleAllGroups(fig) {
  const body = document.getElementById(fig===3 ? 'fig3body' : 'fig4body');
  const set = fig===3 ? S.f3g : S.f4g;
  const els = body ? [...body.querySelectorAll('[data-gk]')] : [];
  const keys = [...new Set(els.map(el=>el.dataset.gk))];
  const anyOpen = keys.some(k => set.has(k));
  if (anyOpen) for (const k of keys) set.delete(k); else for (const k of keys) set.add(k);
  refreshFig(fig);
}

function renderFig3(list) {
  const big = list.filter(m => m.size > 50 && (gateVol(m) || (isNew(m)&&newGate(m))) && S.sz3.has(szOf3(m)));
  const dim = S.fig3dim;                                    // 'refresh' | 'mem'
  const colKeyOf = dim==='mem'
    ? (m => m.mem_band || '待补')
    : (m => hzBandOf(m.refresh_hz));
  const g = new Map();
  for (const m of big) {
    const k = bandVal(m,'part_band') + '||' + colKeyOf(m);
    if (!g.has(k)) g.set(k, []);
    g.get(k).push(m);
  }
  const rows = D.partOrder.filter(p => big.some(m => bandVal(m,'part_band')===p));
  if (big.some(m => bandVal(m,'part_band')==='待补')) rows.push('待补');
  if (!rows.length) rows.push('待补');
  let cols;
  if (dim === 'mem') {
    const present = [...new Set(big.map(colKeyOf))];
    cols = D.memOrder.filter(c => present.includes(c));
    if (big.some(m => colKeyOf(m) === '待补')) cols.push('待补');
  } else {
    const present = [...new Set(big.map(colKeyOf).filter(c => c !== '待补'))];
    const rank = s => s==='待补' ? Infinity : (parseInt(s,10)||0);
    cols = present.sort((a,b) => rank(a)-rank(b));
    if (big.some(m => colKeyOf(m) === '待补')) cols.push('待补');
  }
  if (!cols.length) cols.push('待补');
  const visRows = rows.filter(r => visibleF(S.f3Rows, r));
  const visCols = cols.filter(c => visibleF(S.f3Cols, c));
  const showAll = (S.f3Rows == null) && (S.f3Cols == null);
  let h = ftHead('控光分区 ＼ ' + (dim==='mem' ? '内存' : '刷新率'),
                 cols.map(c => `<span class="fbtn col ${visibleF(S.f3Cols,c)?'on':'off'}" data-f3c="${c}">${c}<br><span class="cnt-label">${big.filter(m=>colKeyOf(m)===c).length}款</span></span>`));
  for (const p of rows) {
    const rowOn = visibleF(S.f3Rows, p);
    h += `<tr class="frow ${rowOn?'':'off'}"><td class="rowhead"><div class="rl"><span class="fbtn row ${rowOn?'on':'off'}" data-f3r="${p}">${partLabel(p)}</span><span class="cnt">${big.filter(m=>bandVal(m,'part_band')===p).length}款</span></div></td>`;
    for (const c of cols) {
      if (!visibleF(S.f3Cols, c)) { h += '<td class="col-off"></td>'; continue; }
      const arr = (g.get(p+'||'+c)||[]).sort((a,b)=>b.cum_vol-a.cum_vol);
      h += `<td><div class="gridbox">${arr.length ? cellGroups(arr, true, p+'||'+c, S.f3g) : '<span class="empty-dim">—</span>'}</div></td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  document.getElementById('fig3body').innerHTML = h;
  const noP = big.length - big.filter(m=>m.part_band).length;
  document.getElementById('fig3cnt').textContent =
    `${big.length} 款 >50吋主力型号 · 参数覆盖 ${big.length-noP} 款` + (noP ? ` · 待补 ${noP} 款` : '');
  const tb = document.getElementById('fig3tools');
  if (tb) tb.innerHTML =
    `<span class="fbtn ${dim==='refresh'?'on':''}" id="f3d_refresh">分区 × 刷新率</span>
     <span class="fbtn ${dim==='mem'?'on':''}" id="f3d_mem">分区 × 内存</span>
     <span class="fbtn all ${showAll?'on':'off'}" id="f3All">↺ 全选 / 重置</span>
     <span class="fbtn all" id="f3exp">全部展开</span>
     <span class="fsum">${showAll ? '全部显示' : '筛选: '+visRows.length+'/'+rows.length+' 行 · '+visCols.length+'/'+cols.length+' 列'}</span>`;
  const body = document.getElementById('fig3body');
  body.querySelectorAll('[data-f3c]').forEach(el => el.onclick = e => { e.stopPropagation(); S.f3Cols = toggleF(S.f3Cols, el.dataset.f3c); refreshFig(3); });
  body.querySelectorAll('[data-f3r]').forEach(el => el.onclick = e => { e.stopPropagation(); S.f3Rows = toggleF(S.f3Rows, el.dataset.f3r); refreshFig(3); });
  const allBtn = document.getElementById('f3All'); if (allBtn) allBtn.onclick = () => { S.f3Rows = null; S.f3Cols = null; refreshFig(3); };
  bindGroupHeads(3);
  const bx = document.getElementById('f3exp'); if (bx) bx.onclick = () => toggleAllGroups(3);
  const dr = document.getElementById('f3d_refresh'); if (dr) dr.onclick = () => { S.fig3dim='refresh'; refreshFig(3); };
  const dm = document.getElementById('f3d_mem');     if (dm) dm.onclick = () => { S.fig3dim='mem'; refreshFig(3); };
}

function renderFig4(list) {
  const sm = list.filter(m => m.size <= 50 && ((m.vol20||0) > 1000 || (isNew(m)&&newGate(m))) && m.size !== 25 && S.sz4.has(szOf4(m)));
  const rowKey = S.fig4 === 'energy' ? 'energy' : 'mem_band';
  const rowOrder = S.fig4 === 'energy' ? D.energyOrder : D.memOrder;
  const colKeyOf = m => m.res_band || '待补';
  const g = new Map();
  for (const m of sm) {
    const k = bandVal(m,rowKey) + '||' + colKeyOf(m);
    if (!g.has(k)) g.set(k, []);
    g.get(k).push(m);
  }
  const rows = rowOrder.filter(r => sm.some(m => bandVal(m,rowKey)===r));
  if (!rows.length) rows.push('待补');
  const colVals = [...new Set(sm.map(colKeyOf))];
  const cols = D.resOrder.filter(c => colVals.includes(c));
  if (!cols.length) cols.push('待补');
  const visRows = rows.filter(r => visibleF(S.f4Rows, r));
  const visCols = cols.filter(c => visibleF(S.f4Cols, c));
  const showAll = (S.f4Rows == null) && (S.f4Cols == null);
  let h = ftHead(S.fig4==='energy' ? '能效 ＼ 分辨率' : '内存 ＼ 分辨率',
                 cols.map(c => `<span class="fbtn col ${visibleF(S.f4Cols,c)?'on':'off'}" data-f4c="${c}">${c}<br><span class="cnt-label">${sm.filter(m=>colKeyOf(m)===c).length}款</span></span>`));
  for (const r of rows) {
    const rowOn = visibleF(S.f4Rows, r);
    h += `<tr class="frow ${rowOn?'':'off'}"><td class="rowhead"><div class="rl"><span class="fbtn row ${rowOn?'on':'off'}" data-f4r="${r}">${r}</span><span class="cnt">${sm.filter(m=>bandVal(m,rowKey)===r).length}款</span></div></td>`;
    for (const c of cols) {
      if (!visibleF(S.f4Cols, c)) { h += '<td class="col-off"></td>'; continue; }
      const arr = (g.get(r+'||'+c)||[]).sort((a,b)=>b.cum_vol-a.cum_vol);
      h += `<td><div class="gridbox">${arr.length ? cellGroups(arr, false, r+'||'+c, S.f4g) : '<span class="empty-dim">—</span>'}</div></td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  document.getElementById('fig4body').innerHTML = h;
  const tb = document.getElementById('fig4tools');
  if (tb) tb.innerHTML = `<span class="fbtn ${S.fig4==='mem'?'on':''}" id="f4d_mem">内存 × 分辨率</span>
    <span class="fbtn ${S.fig4==='energy'?'on':''}" id="f4d_energy">能效 × 分辨率</span>
    <span class="fbtn all ${showAll?'on':'off'}" id="f4All">↺ 全选 / 重置</span>
    <span class="fbtn all" id="f4exp">全部展开</span>
    <span class="fsum">${showAll ? '全部显示' : '筛选: '+visRows.length+'/'+rows.length+' 行 · '+visCols.length+'/'+cols.length+' 列'}</span>`;
  const body = document.getElementById('fig4body');
  body.querySelectorAll('[data-f4c]').forEach(el => el.onclick = e => { e.stopPropagation(); S.f4Cols = toggleF(S.f4Cols, el.dataset.f4c); refreshFig(4); });
  body.querySelectorAll('[data-f4r]').forEach(el => el.onclick = e => { e.stopPropagation(); S.f4Rows = toggleF(S.f4Rows, el.dataset.f4r); refreshFig(4); });
  const allBtn = document.getElementById('f4All'); if (allBtn) allBtn.onclick = () => { S.f4Rows = null; S.f4Cols = null; refreshFig(4); };
  const dm = document.getElementById('f4d_mem'); if (dm) dm.onclick = () => { S.fig4='mem'; refreshFig(4); };
  const de = document.getElementById('f4d_energy'); if (de) de.onclick = () => { S.fig4='energy'; refreshFig(4); };
  bindGroupHeads(4);
  const bx = document.getElementById('f4exp'); if (bx) bx.onclick = () => toggleAllGroups(4);
  document.getElementById('fig4cnt').textContent = `${sm.length} 款 ≤50吋（>1000 台）`;
  // 待补提示（动态真实覆盖，数据补齐后自动消失）：能效/内存视图分别统计
  const warnEl = document.getElementById('fig4warn');
  const dimLab = S.fig4 === 'energy' ? '能效等级' : '内存';
  const noN = sm.filter(m => bandVal(m, rowKey) === '待补').length;
  if (noN > 0) {
    warnEl.innerHTML = `<b>${dimLab}待补 ${noN} / ${sm.length} 款：</b>这些型号暂无${dimLab}数据源，以「待补」行/列占位；本地参数表补齐后重新生成即自动归位。`;
    warnEl.style.display = '';
  } else { warnEl.style.display = 'none'; }
}

function renderAll() {
  const base = filtered();     // 头部统计用全局口径（品牌无关）
  document.getElementById('statN').textContent = base.length;
  document.getElementById('statV').textContent = fmt(base.reduce((s,m)=>s+m.cum_vol,0));
  fillNewSec();
  renderFig1(figList(1)); renderFig2(figList(2)); renderFig3(figList(3)); renderFig4(figList(4));
  buildFigBrands(1); buildFigBrands(2); buildFigBrands(3); buildFigBrands(4);
  [1,2,3,4].forEach(f => buildClearCtl(f, 'fig'+f+'brand'));
  syncTableWidth();
  bindCells();
  buildFloatHeads();
  saveHash();
}

function bindCells() {
  document.querySelectorAll('.cell[data-series]').forEach(el => {
    el.onclick = () => {
      if (S.cmp) return;
      const k = el.dataset.brand+'|'+el.dataset.series+'|'+el.dataset.col;
      if (S.open[k]) { delete S.open[k]; } else { S.open[k] = 1; }
      drawExp();
    };
  });
  document.querySelectorAll('.cell[data-belt]').forEach(el => {
    el.onclick = () => {
      if (S.cmp) return;
      const k = 'B|'+el.dataset.belt+'|'+el.dataset.size;
      if (S.open[k]) { delete S.open[k]; } else { S.open[k] = 1; }
      drawExp();
    };
  });
  document.querySelectorAll('.mc').forEach(el => {
    el.onclick = e => {
      e.stopPropagation();
      if (!S.cmp) return;
      const id = +el.dataset.id;
      const i = S.sel.indexOf(id);
      if (i >= 0) S.sel.splice(i,1);
      else { if (S.sel.length >= 8) { alert('最多对比 8 个型号'); return; } S.sel.push(id); }
      renderAll(); renderCmpBar();
    };
  });
}

function drawExp() {
  document.getElementById('expbox').innerHTML = '';
  const box = document.getElementById('expbox');
  for (const k in S.open) {
    const [a,b,c] = k.split('|');
    let arr;
    if (a === 'B') arr = figList(2).filter(m => m.fig1_ok && beltOf(m)===b && fig2ColOf(m)===c);
    else arr = figList(1).filter(m => m.fig1_ok && m.brand===a && m.series===b && String(m.fig1_col)===c);
    arr.sort((x,y)=>y.cum_vol-x.cum_vol);
    // 品牌筛选 chips（图2 方案 A）：多选，只过滤本格展开列表；点满取消即恢复全部
    const brands = [...new Set(arr.map(m=>m.brand))];
    const fset = S.expBrands[k];
    const vis = fset ? arr.filter(m => fset.has(m.brand)) : arr;
    const chipsHtml = brands.length > 1
      ? `<div class="exp-brands">` + brands.map(bn =>
          `<span class="ebtn ${(!fset || fset.has(bn))?'':'off'}" data-expk="${k}" data-eb="${bn}" style="background:${CO[bn]||'#8b95a3'}">${bn}</span>`).join('') + `</div>`
      : '';
    const d = document.createElement('div');
    d.className = 'exp';
    const cLab = a === 'B' ? c : (D.fig1Lab[String(c)] || (c+'吋'));
    d.innerHTML = `<div class="exp-h">${a==='B'?b+' 价格带':a+' '+b+' 系列'} · ${cLab} · ${vis.length}${fset?'/'+arr.length:''} 款 <span class="exp-fold" data-expk="${k}" title="收起本 Panel">收起 ▴</span></div>`
      + chipsHtml + `<div class="gridbox">${vis.map(m=>chip(m,true)).join('')}</div>`;
    box.appendChild(d);
  }
  box.querySelectorAll('.ebtn').forEach(el => el.onclick = e => {
    e.stopPropagation();
    const k2 = el.dataset.expk, bn = el.dataset.eb;
    S.expBrands[k2] = toggleF(S.expBrands[k2], bn);
    drawExp();
  });
  box.querySelectorAll('.exp-fold').forEach(el => el.onclick = e => {
    e.stopPropagation();
    delete S.open[el.dataset.expk];
    drawExp();
    saveHash();
  });
  bindCells();
}

/* ===== 图1/图2 一键展开 / 收起全部 ===== */
const EXPKEYS = {1: [], 2: []};
function updateExpBtn(fig, keys) {
  EXPKEYS[fig] = keys || [];
  const btn = document.getElementById(fig === 1 ? 'fig1exp' : 'fig2exp');
  if (!btn) return;
  const n = EXPKEYS[fig].filter(k => S.open[k]).length;
  btn.textContent = n ? `⊖ 收起全部（已展开 ${n}）` : `⊕ 展开全部（${EXPKEYS[fig].length} 格）`;
}
function toggleAllExp(fig) {
  const keys = EXPKEYS[fig];
  if (!keys.length) return;
  const anyOpen = keys.some(k => S.open[k]);
  if (anyOpen) {
    for (const k of keys) delete S.open[k];
  } else {
    if (keys.length > 60 && !confirm(`将展开 ${keys.length} 个格子的型号明细，页面会明显变长，确认继续？`)) return;
    for (const k of keys) S.open[k] = 1;
  }
  drawExp();
  updateExpBtn(fig, keys);
}

/* ===== 导出当前筛选结果 CSV（带 BOM，Excel 可直接打开不乱码） ===== */
function exportCsv() {
  const list = filtered();
  if (!list.length) { alert('当前筛选条件下没有型号可导出'); return; }
  const cols = ['品牌','系列','型号','产品名','尺寸吋','累计销量','均价','价格带','屏幕技术','音响','控光分区','刷新率Hz','内存','抗反射','分辨率','能效等级'];
  // 中文列名 → 模型英文字段（修复：此前用中文 key 取数导致品牌等列全空）
  const F = {'品牌':'brand','系列':'series','型号':'model','产品名':'product_name','价格带':'price_belt',
    '屏幕技术':'tech','音响':'audio','控光分区':'part','刷新率Hz':'refresh_hz','内存':'mem',
    '抗反射':'anti','分辨率':'res','能效等级':'energy'};
  const val = (m, k) => {
    if (k === '尺寸吋') return m.size;
    if (k === '累计销量') return m.cum_vol;
    if (k === '均价') return m.avg_price;
    if (k === '音响') return m.audio || '待补';   // 音响维度暂无数据源，占位
    const fk = F[k];
    if (fk) return m[fk] ?? '';
    return m[k] ?? '';
  };
  const esc = v => { const s = String(v == null ? '' : v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const rows = [cols.join(',')];
  for (const m of list) rows.push(cols.map(k => esc(val(m, k))).join(','));
  const blob = new Blob(['\uFEFF' + rows.join('\r\n')], {type: 'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  const ts = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  a.href = URL.createObjectURL(blob);
  a.download = `8品牌AVC筛选结果_${list.length}款_${ts}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 3000);
}

/* ===== 图5 对标表：下载竞争地图选中型号参数 CSV ===== */
function f5TechPlain(m) {
  const n = f5TechNorm(m && m.tech ? String(m.tech).trim() : '');
  return n || '待补';
}
function f5MemPlain(mem) {
  if (!mem) return '待补';
  return String(mem).replace(/\b(\d+)\.0(?=[+×Xx]|$)/g, '$1');
}
function exportF5Csv(list) {
  if (!list.length) { alert('当前选中格子内没有可导出的型号'); return; }
  const ts = F5._tiers || {has:false, kmeans:[]};
  const cols = ['品牌','型号/产品名','尺寸吋','累计销量','均价','配置分','配置档','屏幕技术',`近8周销量(26W30-37)`,'原生刷新率Hz','控光分区','内存','抗反射','音响'];
  const val = (m, k) => {
    const w = ((m.gro||{}).w8||[]).reduce((s,v)=>s+v,0);
    const hits = {
      '品牌': m.brand,
      '型号/产品名': shortName(m),
      '尺寸吋': m.size,
      '累计销量': m.cum_vol,
      '均价': m.avg_price,
      '配置分': f5NoScore(m) ? '待补' : f5Score(m) + '/100',
      '配置档': f5NoScore(m) ? '待补' : f5TierOf(f5Score(m), ts).label,
      '屏幕技术': f5TechPlain(m),
      '近8周销量(26W30-37)': w,
      '原生刷新率Hz': m.refresh_hz != null && String(m.refresh_hz).trim()!=='' ? String(m.refresh_hz).replace(/Hz$/i,'') + 'Hz' : '待补',
      '控光分区': m.part != null && m.part!=='' ? m.part : (m.part_band ? (m.part_band==='无分区' ? '无' : m.part_band) : '待补'),
      '内存': f5MemPlain(m.mem),
      '抗反射': (m.anti||m.anti_band) || '待补',
      '音响': m.audio || '待补',
    };
    return hits[k];
  };
  const esc = v => { const s = String(v == null ? '' : v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const rows = [cols.join(',')];
  rows.push(['—','—','—','—','—','—','—','—','—','—','—','—','—','—'].join(','));
  for (const m of list) rows.push(cols.map(k => esc(val(m, k))).join(','));
  rows.push(['','','','','合计','', list.reduce((s,m)=>s+m.cum_vol,0),'','','','','',''].join(','));
  const blob = new Blob(['\uFEFF' + rows.join('\r\n')], {type: 'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  const tsd = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  a.href = URL.createObjectURL(blob);
  a.download = `图5对标_${list.length}款_${tsd}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 3000);
}

/* ===== URL 状态记忆：筛选写入 hash（replaceState 不产生历史），刷新/分享不丢 ===== */
function saveHash() {
  const o = {b1:[...S.b[1]], b2:[...S.b[2]], b3:[...S.b[3]], b4:[...S.b[4]],
             s: S.size, bl: [...S.belts], q: S.q, mo: S.mainOnly ? 1 : 0, f4: S.fig4};
  const h = '#' + encodeURIComponent(JSON.stringify(o));
  if (location.hash !== h) history.replaceState(null, '', h);
}
function loadHash() {
  if (!location.hash || location.hash.length < 3) return;
  try {
    const o = JSON.parse(decodeURIComponent(location.hash.slice(1)));
    if (Array.isArray(o.b1)) S.b[1] = new Set(o.b1.filter(b => D.brands.includes(b)));
    if (Array.isArray(o.b2)) S.b[2] = new Set(o.b2.filter(b => D.brands.includes(b)));
    if (Array.isArray(o.b3)) S.b[3] = new Set(o.b3.filter(b => D.brands.includes(b)));
    if (Array.isArray(o.b4)) S.b[4] = new Set(o.b4.filter(b => D.brands.includes(b)));
    if (typeof o.s === 'string' && ['全部'].concat(D.sizeOrder).includes(o.s)) S.size = o.s;
    if (Array.isArray(o.bl)) S.belts = new Set(o.bl.filter(x => D.beltOrder.concat(['待补']).includes(x)));
    if (typeof o.q === 'string') S.q = o.q;
    if (o.mo) S.mainOnly = true;
    if (o.f4 === 'mem' || o.f4 === 'energy') S.fig4 = o.f4;
  } catch (e) { /* 非法 hash 忽略，按默认状态渲染 */ }
}

// ---- 每图独立品牌筛选（多选 chips + 全选/清空），只作用于本图 ----
function buildFigBrands(fig) {
  const el = document.getElementById('fig'+fig+'brand');
  if (!el) return;
  const bs = S.b[fig];
  el.innerHTML = D.brands.map(b => {
    const on = bs.has(b);
    return `<span class="chip ${on?'on':''}" data-fb="${fig}" data-b="${b}"${on?` style="background:${CO[b]};border-color:${CO[b]};color:#fff"`:''}><span class="dot"></span>${b}</span>`;
  }).join('') + `<span class="btn mini" data-fba="${fig}">全选</span><span class="btn mini" data-fbn="${fig}">清空</span>`;
  el.querySelectorAll('.chip[data-fb]').forEach(c => c.onclick = () => {
    const b = c.dataset.b;
    if (bs.has(b)) { bs.delete(b); c.classList.remove('on'); c.style.background='#fff'; c.style.color=''; }
    else { bs.add(b); c.classList.add('on'); c.style.background=CO[b]; c.style.color='#fff'; }
    refreshFig(fig);
  });
  el.querySelectorAll('[data-fba]').forEach(x => x.onclick = () => {
    bs.clear(); D.brands.forEach(b => bs.add(b));
    el.querySelectorAll('.chip[data-fb]').forEach(c => { c.classList.add('on'); c.style.background=CO[c.dataset.b]; c.style.color='#fff'; });
    refreshFig(fig);
  });
  el.querySelectorAll('[data-fbn]').forEach(x => x.onclick = () => {
    bs.clear();
    el.querySelectorAll('.chip[data-fb]').forEach(c => { c.classList.remove('on'); c.style.background='#fff'; c.style.color=''; });
    refreshFig(fig);
  });
}

function buildSizes() {
  const el = document.getElementById('crowSizes');
  if (!el) return;
  const arr = ['全部'].concat(D.sizeOrder);
  el.innerHTML = arr.map(s => `<span class="btn ${s===S.size?'on':''}" data-s="${s}">${s}</span>`).join('');
  el.querySelectorAll('.btn').forEach(b => b.onclick = () => {
    el.querySelectorAll('.btn').forEach(x => x.classList.remove('on'));
    b.classList.add('on'); S.size = b.dataset.s; renderAll();
  });
}

/* 图3/图4 独立尺寸筛选（多选 chips + 全部按钮），只作用于本图 */
function buildFigSizes(fig) {
  const el = document.getElementById('fig'+fig+'size');
  if (!el) return;
  const order = fig === 3 ? SZ3 : SZ4;
  const set = fig === 3 ? S.sz3 : S.sz4;
  el.innerHTML = `<span class="slab">尺寸：</span>` + order.map(b =>
    `<span class="chip ${set.has(b)?'on':''}" data-fs="${fig}" data-z="${b}">${b}</span>`).join('') +
    `<span class="btn mini" data-fszall="${fig}">全部</span>`;
  el.querySelectorAll('.chip[data-fs]').forEach(c => c.onclick = () => {
    const b = c.dataset.z;
    if (set.has(b)) { set.delete(b); c.classList.remove('on'); }
    else { set.add(b); c.classList.add('on'); }
    refreshFig(fig);
  });
  el.querySelectorAll('[data-fszall]').forEach(x => x.onclick = () => {
    set.clear(); order.forEach(b => set.add(b));
    el.querySelectorAll('.chip[data-fs]').forEach(c => c.classList.add('on'));
    refreshFig(fig);
  });
}

function buildMisc() {
  const qEl = document.getElementById('q');
  qEl.value = S.q;                              // URL hash 恢复的搜索词回填
  qEl.oninput = e => {
    S.q = e.target.value.trim();
    if (f5RelCfg) f5RelCfg._base = null;
    renderAll();
    try {                                    // 图5 独立渲染，renderAll 不重绘 → 此处联动锁品牌/跳格子
      const ck = f5SearchApply();
      f5BrandFilter(); f5Grid(); f5Draw(); f5SyncPanel();
      if (ck && ck !== 'RESET') {
        const el = document.querySelector('#f5map .f5cell[data-k="' + ck + '"]');
        if (el) el.scrollIntoView({ block: 'center', inline: 'center', behavior: 'smooth' });
      }
    } catch (e) {}
  };
  document.getElementById('csvBtn').onclick = exportCsv;
  document.getElementById('fig2exp').onclick = () => toggleAllExp(2);
  document.getElementById('cmpBtn').onclick = e => {
    S.cmp = !S.cmp;
    e.target.classList.toggle('on', S.cmp);
    document.getElementById('cmpBar').classList.toggle('show', S.cmp);
    document.body.classList.toggle('cmpOn', S.cmp);
    if (!S.cmp) { S.sel = []; renderCmpBar(); }
    renderAll();
  };
  // Esc 关闭对比弹窗
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      document.getElementById('cmpModal').classList.remove('show');
    }
  });
}

function renderCmpBar() {
  const bar = document.getElementById('cmpList');
  bar.innerHTML = S.sel.map(id => {
    const m = M[id], c = CO[m.brand];
    return `<span class="cb-i" style="border-left-color:${c}"><span class="md" style="width:7px;height:7px;border-radius:50%;background:${c};display:inline-block"></span>${m.brand} ${m.model} ${m.size}吋<span class="x" data-x="${id}">✕</span></span>`;
  }).join('');
  bar.querySelectorAll('.x').forEach(x => x.onclick = () => {
    S.sel = S.sel.filter(i => i !== +x.dataset.x); renderAll(); renderCmpBar();
  });
  document.getElementById('cmpN').textContent = S.sel.length;
  document.getElementById('goCmp').disabled = S.sel.length < 2;
}

// 统一换算为「越大越好」的可比序值；返回 null 表示不参与最优/最差判定
function rankIdx(k, m) {
  if (k === 'part') { const i = D.partOrder.indexOf(m.part_band || '待补'); return i < 0 ? null : D.partOrder.length - i; }
  if (k === 'mem')  { const i = D.memOrder.indexOf(m.mem_band || '待补');  return i < 0 ? null : i; }
  if (k === 'refresh') return m.refresh_hz || null;
  if (k === 'energy') { const i = D.energyOrder.indexOf(m.energy || '待补'); return i < 0 ? null : D.energyOrder.length - i; }
  if (k === 'cum_vol') return m.cum_vol || null;
  return null;
}

function doCompare() {
  const sel = S.sel.map(id => M[id]);
  const keys = ['size','cum_vol','avg_price','price_belt','tech']
    .concat(sel.some(m=>m.size>50) ? ['part','refresh','mem','anti'] : [])
    .concat(sel.some(m=>m.size<=50) ? ['res','mem','energy'] : []);
  const seen = new Set(); const ks = keys.filter(k => !seen.has(k) && seen.add(k));
  const LAB = {size:'尺寸', cum_vol:'累计销量', avg_price:'均价', price_belt:'价格带', tech:'屏幕技术',
    part:'控光分区', refresh:'刷新率', mem:'内存', anti:'抗反射', res:'分辨率', energy:'能效等级'};
  const noColor = new Set(['avg_price','price_belt','size']);
  let h = '<table class="ctab"><tr><th>指标</th>' + sel.map(m =>
    `<th style="border-top:3px solid ${CO[m.brand]}">${m.brand}<br>${m.model}<br><span style="font-weight:400;color:#8b95a3">${m.size}吋</span></th>`).join('') + '</tr>';
  for (const k of ks) {
    const vals = sel.map(m => {
      if (k === 'size') return m.size;
      if (k === 'cum_vol') return m.cum_vol;
      if (k === 'avg_price') return m.avg_price;
      if (k === 'price_belt') return m.price_belt || '待补';
      return m[k] || null;
    });
    h += `<tr><td class="lab">${LAB[k]}</td>`;
    let bi = -1, wi = -1;
    if (!noColor.has(k)) {
      const idx = sel.map((m,i) => ({i, r:rankIdx(k, m)})).filter(x => x.r != null);
      if (idx.length > 1) {
        idx.sort((a,b) => b.r - a.r);
        if (idx[0].r !== idx[idx.length-1].r) { bi = idx[0].i; wi = idx[idx.length-1].i; }
      }
    }
    vals.forEach((v,i) => {
      let cls = '';
      if (!v && v !== 0) cls = 'na';
      else if (i === bi) cls = 'best';
      else if (i === wi) cls = 'worst';
      else if (noColor.has(k)) cls = '';
      let txt;
      if (cls === 'na') txt = '待补';
      else if (k === 'cum_vol') txt = fmt(v)+'台';
      else if (k === 'avg_price') txt = '¥'+v;
      else if (k === 'size') txt = v+'吋';
      else txt = v;
      h += `<td class="${cls}">${txt}</td>`;
    });
    h += '</tr>';
  }
  h += '</table>';
  document.getElementById('cmpBody').innerHTML = h;
  document.getElementById('cmpModal').classList.add('show');
}

document.getElementById('goCmp').onclick = doCompare;
document.getElementById('cmpClear').onclick = () => { S.sel = []; renderAll(); renderCmpBar(); };
document.getElementById('cmpExit').onclick = () => {
  S.cmp = false; S.sel = [];
  document.getElementById('cmpBtn').classList.remove('on');
  document.getElementById('cmpBar').classList.remove('show');
  document.body.classList.remove('cmpOn');
  renderCmpBar(); renderAll();
};
document.getElementById('mClose').onclick = () => document.getElementById('cmpModal').classList.remove('show');
document.getElementById('cmpModal').onclick = e => { if (e.target.id === 'cmpModal') e.currentTarget.classList.remove('show'); };
document.getElementById('diffOnly').onchange = e => {
  const on = e.target.checked;
  document.querySelectorAll('#cmpBody tr').forEach((tr,i) => {
    if (!i) return;
    const tds = [...tr.querySelectorAll('td')].slice(1);
    const set = new Set(tds.map(t => t.textContent));
    tr.style.display = (on && set.size <= 1) ? 'none' : '';
  });
};

loadHash(); buildMisc(); buildFigSizes(3); buildFigSizes(4); renderAll();

/* ============================ 图5 单产品定位分析 ============================ */
// 尺寸档位（仅 >50 吋）
const F5SZ = [['55','55'],['65','65'],['75','75'],['85','85'],['98+','98+']];
const F5SZ_OF = m => { const s = m.size||0; return s<=60?'55':s<=69?'65':s<=79?'75':s<=92?'85':'98+'; };
// 价位段：12 档（按每款型号均价 avg_price 精确定档；3000-4999 拆 3000-3999/4000-4999，5000-6999 拆 5000-5999/6000-6999）
const F5PR = ['<1999','2000-2999','3000-3999','4000-4999','5000-5999','6000-6999','7000-7999','8000-8999','9000-9999','10000-12999','13000-15999','16000+'];
// 按均价落档；均价缺失/为 0 返回 null（不归类）
const F5PR_OF = m => {
  const a = m.avg_price;
  if (!(a > 0)) return null;
  if (a >= 16000) return '16000+';
  if (a >= 13000) return '13000-15999';
  if (a >= 10000) return '10000-12999';
  if (a >= 9000) return '9000-9999';
  if (a >= 8000) return '8000-8999';
  if (a >= 7000) return '7000-7999';
  if (a >= 6000) return '6000-6999';
  if (a >= 5000) return '5000-5999';
  if (a >= 4000) return '4000-4999';
  if (a >= 3000) return '3000-3999';
  if (a >= 1999) return '2000-2999';
  return '<1999';
};
// 图5 门槛：仅保留 >50 吋；近20周 98吋+(≥98) >200 台，其他 >1000 台。新品豁免此门槛（用统一新品门槛）
const F5_GT = m => { const s = m.size||0, v20 = m.vol20||0; return s >= 98 ? (v20 > 200) : (v20 > 1000); };
// 近8周0销量排除（持续生效）：近8周 w8 (26W30-37) 全部为 0 的型号从图5移除（新品豁免，改用近4周在售判断）
const f5ZeroSales = m => { const w=(m.gro||{}).w8||[]; return w.length && w.every(v=>!v); };
const F5POOL = M.filter(m => (m.size||0) > 50 && (F5_GT(m) || (isNew(m)&&newGate(m)) || f5Guard(m)) && F5PR_OF(m) && (!f5ZeroSales(m) || (isNew(m)&&newActive(m)) || f5Guard(m)));   // 图5 型号池：均价无法归档或近8周0销量不入池；新品与保护期豁免门槛与0销量
const F5 = {
  b: new Set(D.brands),                 // 图5 独立品牌筛选
  sz: null,                             // 目标产品尺寸档（展示用，多规格时置 ''）
  pSet: new Set(),                      // 已选价位段（展示用，取已选格子的并集）
  cells: new Set(),                     // 已选格子集合（"尺寸||价位"），地图多选的真实状态
  cfg: 0,                               // 配置定位：1=高配 0=标配 -1=低价
  filter: {},                           // 列筛选 {tech:Set, part:Set, refresh:Set, anti:Set, mem:Set, audio:Set}
  sort: 'cum_vol', asc: false,   // 默认按累计销量降序
};
// 搜索框联动图5：命中品牌→锁品牌(仅保留该品牌)；命中系列/型号→跳转到销量最高/该型号所在竞争地图格子
function f5SearchApply() {
  const q = (S.q || '').trim().toLowerCase();
  if (!q) { F5.b = new Set(D.brands); return 'RESET'; }   // 清空搜索 → 解除品牌锁定，恢复全品牌
  let brand = '';
  for (const b of D.brands) {                       // 最长品牌前缀匹配
    const bl = b.toLowerCase();
    if (q.indexOf(bl) === 0 && bl.length > brand.length) brand = b;
  }
  if (!brand) return null;                          // 未命中品牌：仅普通过滤，不联动图5
  F5.b = new Set([brand]);                          // 清空其他品牌，仅保留此品牌
  const rem = q.slice(brand.length).trim();         // 剥离品牌，用剩余词匹配系列/型号/产品名
  if (!rem) return null;                            // 品牌级：仅锁品牌，不跳格
  const fields = m => [m.brand, m.series, m.model, m.product_name, m.py].map(v => String(v || '').toLowerCase());
  const cands = F5POOL.filter(m => m.brand === brand && fields(m).some(v => v.includes(rem)));
  if (!cands.length) return null;                   // 只锁到品牌、无可定位型号
  const bestOf = arr => arr.reduce((a, b) => ((b.cum_vol || 0) > (a.cum_vol || 0) ? b : a));
  const exact = cands.filter(m => String(m.model || '').toLowerCase().indexOf(rem) === 0);
  const tgt = exact.length ? bestOf(exact) : bestOf(cands);   // 型号前缀精确落位，否则系列/型号最高销量
  const ck = F5SZ_OF(tgt) + '||' + F5PR_OF(tgt);    // 尺寸档||价位段
  F5.cells = new Set([ck]);
  F5.sz = ck.split('||')[0];
  F5.pSet = new Set([ck.split('||')[1]]);
  return ck;
}
// 由 F5.cells(多选格子) 反推面板展示：单一尺寸则高亮该尺寸，多尺寸置 ''；价位显示并集
function f5SyncPanel() {
  const sizes = new Set(), prices = new Set();
  F5.cells.forEach(c => { const sp = c.split('||'); sizes.add(sp[0]); prices.add(sp[1]); });
  F5.sz = sizes.size === 1 ? [...sizes][0] : '';
  F5.pSet = prices;
  const szEl = document.getElementById('f5sz');
  if (szEl) szEl.querySelectorAll('span').forEach(x => x.classList.toggle('on', x.dataset.v === F5.sz));
  const prEl = document.getElementById('f5pr');
  if (prEl) prEl.querySelectorAll('span').forEach(x => x.classList.toggle('on', F5.pSet.has(x.dataset.v)));
}
function f5pct(x) {
  if (x == null) return null;
  const p = Math.round(x * 100);
  return (p>=0?'+':'') + p + '%';
}
function f5BrandFilter() {
  const el = document.getElementById('fig5brand');
  if (!el) return;
  const bs = F5.b;
  el.innerHTML = D.brands.map(b => {
    const on = bs.has(b);
    return `<span class="chip ${on?'on':''}" data-b5="${b}"${on?` style="background:${CO[b]};border-color:${CO[b]};color:#fff"`:''}><span class="dot"></span>${b}</span>`;
  }).join('') + `<span class="btn mini" data-b5a>全选</span><span class="btn mini" data-b5c>清空</span>`;
  el.querySelectorAll('.chip[data-b5]').forEach(c => c.onclick = () => {
    const b = c.dataset.b5;
    if (bs.has(b)) { bs.delete(b); c.classList.remove('on'); c.style.background='#fff'; c.style.color=''; }
    else { bs.add(b); c.classList.add('on'); c.style.background=CO[b]; c.style.color='#fff'; }
    f5Draw(); f5Grid();
  });
  el.querySelector('[data-b5a]').onclick = () => { bs.clear(); D.brands.forEach(b=>bs.add(b)); f5BrandFilter(); f5Draw(); f5Grid(); };
  el.querySelector('[data-b5c]').onclick = () => { bs.clear(); f5BrandFilter(); f5Draw(); f5Grid(); };
}
function f5TargetChips() {
  const sz = document.getElementById('f5sz'), pr = document.getElementById('f5pr');
  sz.innerHTML = F5SZ.map(([k,v]) => `<span data-v="${v}">${k}</span>`).join('');
  sz.querySelectorAll('span').forEach(s => s.onclick = () => {
    sz.querySelectorAll('span').forEach(x=>x.classList.remove('on')); s.classList.add('on');
    F5.sz = s.dataset.v; if (!F5.pSet.size) F5.pSet.add(F5PR[1]);
    F5.cells = new Set([...F5.pSet].map(p=>F5.sz+'||'+p));   // 单选尺寸 → 重设为该尺寸×已选价位
    f5RelCfg._base = null;
    f5Grid(); f5Draw();
  });
  pr.innerHTML = F5PR.map(v => `<span data-v="${v}" class="${F5.pSet.has(v)?'on':''}">${v}</span>`).join('');
  pr.querySelectorAll('span').forEach(s => s.onclick = () => {
    const v = s.dataset.v;
    const base = F5.sz || (F5.cells.size ? [...F5.cells][0].split('||')[0] : F5PR[1]);  // 多尺寸时回退到首个格子尺寸
    if (F5.pSet.has(v)) { if (F5.pSet.size > 1) { F5.pSet.delete(v); s.classList.remove('on'); } }  // 禁止取消最后一个价位
    else if (F5.pSet.size < 4) { F5.pSet.add(v); s.classList.add('on'); }
    F5.cells = new Set([...F5.pSet].map(p=>base+'||'+p));   // 价位调整 → 重设为单尺寸×新价位
    f5SyncPanel();
    f5RelCfg._base = null;
    f5Grid(); f5Draw();
  });
  const f5cfgLb = {'1':'高配','0':'标配','-1':'低价'};
  const cfgEl = document.getElementById('f5cfg');
  cfgEl.querySelectorAll('span').forEach(s => s.onclick = () => {
    if (!F5.filter['score'] || F5.filter['score'].size === 0) F5.filter['score'] = new Set(['高配','标配','低价']);
    const st = F5.filter['score'];
    const lb = f5cfgLb[s.dataset.v];
    if (st.has(lb)) st.delete(lb); else st.add(lb);
    f5Draw();
  });
  // 默认选取数据量最大的 (尺寸) + 数据量最大的单个价位
  const cntS = new Map(), cntP = new Map();
  F5POOL.forEach(m => { if (F5.b.has(m.brand)) {
    const s = F5SZ_OF(m); cntS.set(s,(cntS.get(s)||0)+1);
    const p = F5PR_OF(m); cntP.set(p,(cntP.get(p)||0)+1);
  }});
  F5.sz = [...cntS.entries()].sort((a,b)=>b[1]-a[1])[0][0];
  F5.pSet = new Set([[...cntP.entries()].sort((a,b)=>b[1]-a[1])[0][0]]);
  F5.cells = new Set([F5.sz+'||'+[...F5.pSet][0]]);
  sz.querySelectorAll('span').forEach(x=>x.classList.toggle('on', x.dataset.v===F5.sz));
  pr.querySelectorAll('span').forEach(x=>x.classList.toggle('on', F5.pSet.has(x.dataset.v)));
}
// 拥挤度自动分桶：以型号数规模做 4 档（空白=0；其余用分位数）
function f5Buckets() {
  const F5VIS = F5POOL.filter(m => isClearShow(5, m) && searchHit(m));
  const ns = [];
  const map = new Map();
  F5VIS.forEach(m => { if (F5.b.has(m.brand)) { const k=F5SZ_OF(m)+'||'+F5PR_OF(m); map.set(k,(map.get(k)||0)+1); } });
  for (const n of map.values()) if (n > 0) ns.push(n);
  ns.sort((a,b)=>a-b);
  const q = p => { const i = Math.round(p*(ns.length-1)); return ns[i]; };
  // 阈值：稀疏 / 中等 / 密集 分界（3.5档之上为稀疏，进入均值附近为中等，高四分位为密集）
  return { map, sparse: Math.max(1, Math.round(q(.30))), dense: Math.max(2, Math.round(q(.80))) };
}
function f5Grid() {
  const buckets = f5Buckets();
  const map = buckets.map;
  const F5VIS = F5POOL.filter(m => isClearShow(5, m) && searchHit(m));
  // 格的销量 + 近8周趋势（格内机型逐周销量汇总，26W30-37）
  const volM = new Map(), fM = new Map(), lM = new Map();
  F5VIS.forEach(m => { if (!F5.b.has(m.brand)) return;
    const k = F5SZ_OF(m)+'||'+F5PR_OF(m);
    map.set(k,(map.get(k)||0));
    volM.set(k,(volM.get(k)||0)+m.cum_vol);
    const w8 = (m.gro||{}).w8 || [];
    if (w8.length) {
      const f = w8.slice(0,4).filter(v=>v>0).reduce((s,x)=>{s+=x;return s},0); // 26W30-33
      const l = w8.slice(4).filter(v=>v>0).reduce((s,x)=>{s+=x;return s},0);   // 26W34-37
      fM.set(k,(fM.get(k)||0)+f);
      lM.set(k,(lM.get(k)||0)+l);
    }
  });
  const growK = new Set();
  for (const k of map.keys()) {
    const f = fM.get(k)||0, l = lM.get(k)||0;
    if ((l+f) >= 500 && f > 0 && (l-f)/f >= 0.20) growK.add(k); // 后4周比前4周销量增长≥20%
  }
  // 渲染
  const mk = document.getElementById('f5map');
  let h = `<div class="f5gh">${F5PR.map(p=>`<span>${p}</span>`).join('')}</div>`;
  F5SZ.forEach(([k,v]) => {
    h += `<div class="f5row"><div class="f5ylab">${k}</div><div class="f5cells">`;
    F5PR.forEach((p,i) => {
      const cellKey = v+'||'+p;
      const n = map.get(cellKey)||0;
      const vol = volM.get(cellKey)||0;
      const grow = growK.has(cellKey);
      const sel = F5.cells.has(v+'||'+p);
      const bg = n===0 ? (sel?'#fef2f2':'#f1f5f9')
        : n<=buckets.sparse ? '#dbeafb'
        : n<=buckets.dense ? '#93c5fd' : '#1d4ed8';
      const txtC = n>buckets.dense ? '#fff' : '#1f2328';
      const gg = grow ? '<span class="gg y">▲</span>' : '';
      h += `<div class="f5cell ${n===0?'gap':''} ${sel?'sel':''}" style="background:${bg};color:${txtC}" data-k="${cellKey}">
        ${n===0 ? '<small>空白</small>' : `<b>${n}</b><small style="color:${txtC}">${fmt(vol)}台</small>`}${gg}</div>`;
    });
    h += `</div></div>`;
  });
  mk.innerHTML = h;
  mk.querySelectorAll('.f5cell').forEach(c => c.onclick = () => {
    const ck = c.dataset.k;
    const sz = ck.split('||')[0];
    // 仅支持同尺寸价格带多选：点到另一尺寸时，清空旧尺寸的所有选中格
    if ([...F5.cells].some(k => k.split('||')[0] !== sz)) F5.cells = new Set();
    if (F5.cells.has(ck)) { if (F5.cells.size > 1) F5.cells.delete(ck); } else F5.cells.add(ck);
    f5SyncPanel();
    f5RelCfg._base = null;
    try { f5Grid(); f5Draw(); } catch (e) { f5Err(e, 'cell'); }
  });
}
// 目标产品配置定位辅助：对该(尺寸×已选价位片)竞品求各列中位，判定每款相对你的档位
function f5RelCfg(m, cfg) {
  const key = F5SZ_OF(m)+'||'+F5PR_OF(m);
  // 预计算片内配置基线（全局缓存一次）
  if (!f5RelCfg._base) f5RelCfg._base = new Map();
  if (!f5RelCfg._base.has(key)) {
    const pool = F5POOL.filter(x => F5.b.has(x.brand) && F5.cells.has(F5SZ_OF(x)+'||'+F5PR_OF(x)) && isClearShow(5, x) && searchHit(x));
    const med = {};
    const ar = k => pool.map(x=>x[k]).filter(v=>v!=null).sort((a,b)=>a-b);
    const partMed = [...pool].map(x=>x.part_band?1:0).reduce((a,b)=>a+b,0)/(pool.length||1);
    // 分区档位量化为数值（百级=1 千级=2 百千级=2.5 2000级=4 三千级=5 四千-七千级=6 八千级+=7）
    const pq = v=>v!=null && !String(v).includes('无') ? {百级:1,几百:1,千级:2,百千级:2.5,'2000级':4,'三千级':5,'四千-七千级':6,'八千级+':7}[String(v)]||1 : 0;
    const ps = pool.map(x=>pq(x.part_band)).filter(v=>v>0).sort((a,b)=>a-b);
    med.part = ps.length? ps[Math.floor(ps.length/2)] : 0;
    const hs = pool.map(x=>x.refresh_hz||0).filter(v=>v>0).sort((a,b)=>a-b);
    med.refresh = hs.length? hs[Math.floor(hs.length/2)] : 0;
    const mrs = pool.map(x=>parseFloat(String(x.mem||'0').split('+')[0])||0).filter(v=>v>0).sort((a,b)=>a-b);
    med.mem = mrs.length? mrs[Math.floor(mrs.length/2)] : 0;
    const aq = v=>{ if(v==null||String(v).includes('免'))return 0; if(String(v).includes('7.1'))return 4; if(String(v).includes('2.1.2'))return 3; if(String(v).includes('2.1'))return 2; if(String(v).includes('2.0'))return 1; return 1; };
    const as = pool.map(x=>aq(x.audio)).filter(v=>v>0).sort((a,b)=>a-b);
    med.audio = as.length? as[Math.floor(as.length/2)] : 0;
    f5RelCfg._base.set(key, med);
  }
  const med = f5RelCfg._base.get(key)||{part:0,refresh:0,mem:0,audio:0};
  // 当前型号相对基线高低分：越级=高(1) 平(0) 低(-1) 按多数列
  const pq = v=>v!=null && !String(v).includes('无') ? {百级:1,几百:1,千级:2,百千级:2.5,'2000级':4,'三千级':5,'四千-七千级':6,'八千级+':7}[String(v)]||1 : 0;
  const aq = v=>{ if(v==null||String(v).includes('免'))return 0; if(String(v).includes('7.1'))return 4; if(String(v).includes('2.1.2'))return 3; if(String(v).includes('2.1'))return 2; if(String(v).includes('2.0'))return 1; return 1; };
  const scores = [
    (pq(m.part_band)||0) - med.part,
    (m.refresh_hz||0) - med.refresh,
    (parseFloat(String(m.mem||'0').split('+')[0])||0) - med.mem,
    aq(m.audio) - med.audio,
  ].filter(s=>Math.abs(s)>0);
  let grade = 0;
  if (scores.length) grade = Math.round(scores.reduce((a,b)=>a+b,0)/scores.length);
  const rel = grade>0?'高':grade<0?'低':'平';
  // 相对你的目标：cfg=1 高配(你>对手) 0 标配(≈) -1 低价(你<对手靠价格)
  let win;
  if (cfg===1) win = rel==='低' ? 'win' : (rel==='高'?'loss':'tie');
  else if (cfg===-1) win = rel==='高' ? 'win' : (rel==='低'?'loss':'tie');
  else win = rel==='平' ? 'win' : (rel==='高'?'loss':'tie');
  return { rel, win };
}
// 分区档位显示名：2000级→两千级；三千级/四千-七千级/八千级+ 直接展示
function partLabel(p) { return ({'2000级':'两千级'})[p] || p; }
// 图5 配置评分（技术35 / 分区25 / 刷新10 / 内存10 / 抗反射10 / 音响10 = 100）=====
const _AUD_BR = ['安桥','帝瓦雷','哈曼','JBL','B&W','Bose','雅马哈','索尼'];
function f5TechScore(t){ const n=f5TechNorm(t); if(!n) return 0;
  return {'OLED':35,'RGB-Mini LED':30,'SQD-Mini LED':28,'BGB-Mini LED':20,'QD-Mini LED':18,'Mini LED':15,'QLED':5,'LED':0}[n] ?? 0; }
function parsePart(v){ if(v==null) return null; const n=Number(String(v).trim()); return isNaN(n)?null:n; }
function f5PartScore(m){
  const raw = parsePart(m && m.part);
  if (raw != null) {
    if (raw>=8000) return 25;
    if (raw>=4000) return 22;
    if (raw>=3000) return 20;
    if (raw>=2000) return 19;
    if (raw>=1500) return 18;
    if (raw>=1000) return 16;
    if (raw>=700) return 12;
    if (raw>=500) return 10;
    if (raw>=200) return 7;
    if (raw>=100) return 5;
    if (raw>=1) return 3;
    return 0;
  }
  const b = String((m && m.part_band) || '');
  if (b.includes('无')) return 0;
  return {'几十区':3,'百级':5,'几百':11,'千级':16,'百千级':16,'2000级':19,'三千级':20,'四千-七千级':22,'八千级+':25}[b]||0;
}
function f5HzScore(h){ if(!h) return 0; if(h>=180) return 10; if(h>=170) return 9; if(h>=165) return 8; if(h>=150) return 7; if(h>=132) return 6; if(h>=120) return 4; return 2; }   // 60Hz=2 / 120Hz=4 / 132-144=6 / 150=7 / 165=8 / 170=9 / 180=10
function f5MemScore(m){
  const s = String(m.mem||'').trim();
  const g = s.match(/(?:^|\D)(\d+(?:\.\d+)?)\s*\+\s*(\d+)/);
  if (g) {
    const ram = parseFloat(g[1]), rom = +g[2];
    // RAM：1=0 / 1.5=2 / 2=4 / 3=5 / ≥4=6
    const rb = ram >= 4 ? 6 : (ram >= 3 ? 5 : (ram >= 2 ? 4 : (ram >= 1.5 ? 2 : 0)));
    // ROM：8=0 / 16=1 / 32=2 / 64=3 / ≥128=4
    const add = rom >= 128 ? 4 : (rom >= 64 ? 3 : (rom >= 32 ? 2 : (rom >= 16 ? 1 : 0)));
    return Math.min(10, rb + add);                 // 1+8=0 / 2+32=6 / 3+64=8 / 4+128=10
  }
  const b = String(m.mem_band || '');
  if (b === '≤2GB') return 4; if (b === '3GB') return 5; if (b === '4GB+') return 6; // 兜底：取各档代表值
  return 0;
}
function f5AntiScore(a){ const s=String(a||''); if(!s||s.includes('无')) return 0;
  if(s.includes('+AG')) return 8;   // LR+AG
  if(s.includes('AG')) return 4;
  return 10;                        // LR 及含 LR 的高级偏光屏
}
function f5AudioScore(a){ const s=String(a||''); if(!s||s.includes('免')||s.includes('待定')) return 0;
  const m=s.match(/(\d)\.(\d)(?:\.(\d+))?/); if(!m) return 0;
  const A=+m[1], B=+m[2], C=m[3]?+m[3]:null;
  const famous=_AUD_BR.some(b=>s.includes(b));
  if(C==null){ // a.b 两声道
    if(A===2 && B===0) return 0;          // 2.0
    return 4;                             // 2.1 / 2.2 ...
  }
  // a.b.c 高度声道：2.1.2 为基准；更高(如 2.2.2 / 7.1.2) 记"以上"
  if(A===2 && B===1 && C===2) return 6;   // 2.1.2
  return famous ? 10 : 8;                 // 2.1.2以上：名品=10，否则8
}
function f5Score(m){ return Math.round(
  f5TechScore(m.tech)+f5PartScore(m)+f5HzScore(m.refresh_hz)+
  f5MemScore(m)+f5AntiScore(m.anti)+f5AudioScore(m.audio)); }
// 格子内按配置分 K-Means 聚 3 簇 → 高配/标配/低价（自然分簇，非均等分位；分簇不明显时按分差就近纳入）
const F5_SCORE_KEYS = ['tech','part_band','refresh_hz','mem_band','anti','audio'];
function f5ParamMissing(m){ return F5_SCORE_KEYS.some(k => { const v=m[k]; return v==null || String(v).trim()==='' || v==='待补'; }); }
// 清库 且 参数待补 → 配置档空出、不参与打分分档（补齐参数后再评）
function f5NoScore(m){ return !!isClear9(m) && f5ParamMissing(m); }
// 1 维 K-Means 聚 k 簇（分数数组），返回降序排列的分界值
function f5KMeans(vals,k){
  const a=[...vals].sort((x,y)=>x-y);
  if(!a.length) return [];
  if(a.length<=k) return a;
  const centers=[];
  for(let i=0;i<k;i++) centers.push(a[Math.floor(i*(a.length-1)/(k-1))]); // 均匀取初始质心
  for(let it=0;it<50;it++){
    const cl=Array.from({length:k},()=>[]);
    a.forEach(x=>{
      let bi=0,bd=Infinity;
      for(let i=0;i<k;i++){ const d=Math.abs(x-centers[i]); if(d<bd){bd=d;bi=i;} }
      cl[bi].push(x);
    });
    const nc=cl.map(g=> g.length? g.reduce((s,v)=>s+v,0)/g.length : centers[cl.indexOf(g)]);
    const conv=nc.every((c,i)=>Math.abs(c-(centers[i]??0))<0.01);
    centers.splice(0,k,...nc);
    if(conv) break;
  }
  return centers.sort((x,y)=>x-y);
}
// 在当前格子的所有可评分型号里，对配置分做 3 簇，返回三簇质心（升序）
function f5Tiers(listm){ 
  const base = (listm||[]).filter(m=>!f5NoScore(m));
  if(!base||base.length<3) return {has:false, kmeans:[]};
  const ss=base.map(m=>f5Score(m)).filter(v=>!isNaN(v));
  const cs=f5KMeans(ss,3).sort((a,b)=>a-b);
  if(cs.length!==3) return {has:false, kmeans:[]};
  return {has:true, kmeans:cs};
}
function f5TierOf(sc,ts){ if(sc==null||!ts||!ts.has) return {label:'待补',cls:''};
  const cs=ts.kmeans;
  // 取离 sc 最近的质心 → 该质心对应档位（低价/标配/高配）
  let bi=0,bd=Infinity;
  for(let i=0;i<3;i++){ const d=Math.abs(sc-cs[i]); if(d<bd){bd=d;bi=i;} }
  return {label:['低价','标配','高配'][bi], cls:['lo','md','hi'][bi]};
}
// 对标表列筛选值（与显示一致）
function f5Fval(k, m) {
  if (k==='tech') return f5Tech(m).replace(/<[^>]+>/g,'');
  if (k==='part') return partLabel(m.part_band || '待补');
  if (k==='refresh') return (m.refresh_hz!=null? f5Hz(m.refresh_hz).replace(/(<[^>]+>)/g,'') : '待补');
  if (k==='anti') return (m.anti||m.anti_band)||'待补';
  if (k==='mem') return memFmt(m.mem).replace(/<[^>]+>/g,'');
  if (k==='audio') return m.audio || '待补';
  if (k==='score') return f5NoScore(m) ? '待补' : f5TierOf(f5Score(m), F5._tiers||{has:false,kmeans:[]}).label;
  return m[k]!=null? String(m[k]) : '待补';
}
// 品牌+型号：小米/红米/雷鸟/Vidda 显示「品牌+产品名」，其余「品牌+型号」
function f5pmLabel(m){
  const b = m.brand;
  if ((b==='小米'||b==='红米'||b==='雷鸟'||b==='Vidda') && m.product_name) return b+' '+m.product_name;
  return b+' '+(m.model||'').replace(/^\d+\s*/, '');
}
// 深度分析：针对当前选中格子，给出主导品牌/主导产品/增长型号/配置解析/配置错开建议
function f5Deep(list, locLb){
  const el = document.getElementById('f5deep');
  if (!el) return;
  const H = [];
  if (!list.length) { el.innerHTML=''; return; }
  const n = list.length;
  const vol = list.reduce((s,m)=>s+m.cum_vol,0);
  const topVol = list.reduce((a,b)=>a.cum_vol>=b.cum_vol?a:b, list[0]);
  H.push(`<div class="dhead">深度分析 · ${locLb}（${n} 款 · 累计 ${fmt(vol)} 台）</div>`);

  // 1) 主导品牌（按销量）
  const bs = {};
  list.forEach(m=>{ bs[m.brand]=(bs[m.brand]||0)+m.cum_vol; });
  const bsA = Object.entries(bs).sort((a,b)=>b[1]-a[1]).slice(0,6);
  const bmax = bsA[0][1]||1;
  H.push(`<div class="dsec"><div class="dt">① 主导品牌（按销量）</div>`);
  H.push('<div class="bwrap">'+bsA.map(b=>{
    const p = Math.round(b[1]/vol*100);
    return `<div class="drow"><span class="dbn">${b[0]}</span><span class="dbar" style="--w:${Math.max(2,(b[1]/bmax)*100)}%"></span><span class="dbv">${fmt(b[1])}台 · ${p}%</span></div>`;
  }).join('')+'</div></div>');

  // 2+4) 主导产品分析（文字，无表格：头部销量占比 + TOP10 配置分配）
  const top3 = list.slice().sort((a,b)=>b.cum_vol-a.cum_vol).slice(0,3);
  const top10 = list.slice().sort((a,b)=>b.cum_vol-a.cum_vol).slice(0,10);
  const dist = fn => Object.entries(top10.reduce((o,m)=>{ const v=fn(m); if(v) o[v]=(o[v]||0)+1; return o; },{}))
    .sort((a,b)=>b[1]-a[1]).slice(0,2).map(e=>`${e[0]}×${e[1]}`).join('、');
  const tW = m=>{ const w=(m.gro||{}).w8||[]; const fW=w.slice(0,4).reduce((a,b)=>a+(b||0),0), lW=w.slice(4).reduce((a,b)=>a+(b||0),0);
    if(!(fW>0&&lW>0)) return '转平'; const r=(lW-fW)/fW; return r>=0.2?('↑'+f5pct(r)):(r<=-0.2?('↓'+f5pct(-r)):'平'); };
  H.push(`<div class="dsec"><div class="dt">② 主导产品分析</div>`);
  H.push(`<div class="gm">头部销量：${top3.map((m,i)=>`${i+1}.${f5pmLabel(m)} ${fmt(m.cum_vol)}台·${Math.round(m.cum_vol/vol*100)}%`).join('　')}。${top3[0]?`榜首近8周${tW(top3[0])}。`:''}</div>`);
  H.push(`<div class="gm">TOP10 配置分配：技术=${dist(m=>m.tech||null)}；刷新率=${dist(m=>m.refresh_hz?m.refresh_hz+'Hz':null)}；分区=${dist(m=>pb(m.part_band))}；抗反射=${dist(m=>m.anti||m.anti_band||null)}；内存=${dist(m=>m.mem||null)}；音响=${dist(m=>m.audio||null)}。</div>`);
  H.push('</div>');

  // 3) 近8周增长型号
  const grows = list.map(m=>{ if(f5Clear(m)) return null;  // 清库型号不进增长表现
    const w=(m.gro||{}).w8||[]; const fW=w.slice(0,4).reduce((a,b)=>a+(b||0),0), lW=w.slice(4).reduce((a,b)=>a+(b||0),0);
    if(!(fW>0 && lW>0 && (lW-fW)/fW>=0.20)) return null;
    return {m, r:(lW-fW)/fW}; }).filter(Boolean).sort((a,b)=>b.r-a.r).slice(0,6);
  H.push(`<div class="dsec"><div class="dt">③ 近8周增长表现（W34-37 vs W30-33 ≥20%）</div>`);
  if (!grows.length) H.push(`<div class="gm">该格近8周无增长型号，切入更偏防守。</div>`);
  else H.push('<div class="gm">'+grows.map(g=>`<span class="gchip">${f5pmLabel(g.m)}<b>↑${f5pct(g.r)}</b></span>`).join('')+'</div>');
  H.push('</div>');

  // 分区桶显示转换
  function pb(v){ return v==null?'待补':(String(v).includes('2000级')?'两千级':v); }
  H.push(`<div class="dsec"><div class="dt">⑤ 配置错开机会与建议</div>`);
  const adv = f5Advice(list);
  H.push(adv.length ? adv.map(a=>`<div class="gm adv">${a}</div>`).join('') : '<div class="gm">格内配置已较同质，错开空间有限，主要靠价位/渠道竞争。</div>');
  H.push('</div>');

  el.innerHTML = H.join('');
}

// 配置错开机会判定：统计格内主流 vs 缺口，生成差异化建议
function f5Advice(list){
  const cnt = (fn,buckets)=>{ const o={}; list.forEach(m=>{ const v=fn(m); if(v==null)return; const b=buckets?buckets(v):v; o[b]=(o[b]||0)+1; }); return o; };
  const techB = v=> v&&(v.includes('Mini')||v.includes('OLED')||v.includes('QLED')||v.includes('量子点'))?'高端':'LED系';
  const T = cnt(m=>m.tech, techB);            // {高端: n, LED系: n}
  const hasHi = (T['高端']||0) >= 1;
  const antiCnt = cnt(m=>(m.anti||m.anti_band||''), v=> v&&!v.includes('无') ? '有抗反射':'无');
  const hasAnti = (antiCnt['有抗反射']||0) >= 1;
  const hzs = list.map(m=>m.refresh_hz||0);
  const maxHz = Math.max(...hzs), minHz = Math.min(...hzs);
  const H=[];
  if(!hasHi){
    const ledN = T['LED系']||list.length;
    H.push(`格内 ${ledN}/${list.length} 款为 LED系，无高端面板——若预算允许上「Mini LED 级」可形成技术代差，是与格内主导产品错开的最直接手段。`);
  }
  if(maxHz<=60) H.push(`刷新率全为 60Hz——补一档 120Hz（成本极低）即可拿到该格唯一的刷新率卖点。`);
  else if(minHz>0 && maxHz<120) H.push(`格内最高仅 ${maxHz}Hz——上一档 120Hz 及以上可错开。`);
  if(!hasAnti) H.push(`格内均未配抗反射（AG/LR）——大屏强光环境感知强，加 AG 屏幕是低成本加分项。`);
  if(list.some(m=>(m.audio||'').includes('2.1')) && !list.some(m=>(m.audio||'').includes('2.1.2'))) H.push(`音响普遍停在 2.1，可补 2.1.2 高度声道拉开质感；若用上安桥/帝瓦雷等名品更佳。`);
  return H.slice(0,3);
}
// 千分位
function fmtN(x) {
  if (x==null) return '<span class="na">待补</span>';
  return Number(x).toLocaleString('en-US');
}
// 近8周微型柱状图：m.gro.w8 (26W30-37)，柱高按片内最大销量归一
const _w8max = (() => { let mx = 1; F5POOL.forEach(m => { const w=(m.gro||{}).w8||[]; w.forEach(v=>{ if(v>mx) mx=v; }); }); return mx; })();
// 高价位例外档：55吋≥5000 / 65吋≥6000 / 75吋≥7000 / 85吋≥9000 / 98吋+≥16000（按均价）
function f5Excp(m){ const s=m.size||0, p=m.avg_price||0; const z=F5SZ_OF(m);
  if(z==='55') return p>=5000; if(z==='65') return p>=6000; if(z==='75') return p>=7000;
  if(z==='85') return p>=9000; return p>=16000; }
// 清库判定（近8周 w8=26W30-37）：合计 <阈值 或 连续≥4周单周 <阈值；例外档阈值10，普通100
function f5Clear(m){
  const w=(m.gro||{}).w8||[]; if(!w.length) return null;
  if(protEx(m)) return null;   // 新品+保护期豁免清库打标
  const z = F5SZ_OF(m);
  const ex = f5Excp(m);
  const sum = w.reduce((a,b)=>a+(b||0),0);
  if (z==='98+' && !ex) {
    // 98吋+ 低价(<16000)：累计<300 且 连续≥4周单周<10 双条件同时命中才判清库
    let run=0, mxr=0;
    for(const v of w){ run = (v<10) ? run+1 : 0; if(run>mxr) mxr=run; }
    if (!(sum < 300 && mxr >= 4)) return null;
    return { reason:`近8周累计${sum}台 <300台 且连续${mxr}周<10台/周`, ex:false, T:300 };
  }
  const T = ex ? 10 : 100;
  let run=0, maxrun=0;
  for(const v of w){ run = (v<T) ? run+1 : 0; if(run>maxrun) maxrun=run; }
  const hitA = maxrun>=4, hitB = sum<T;
  if(!hitA && !hitB) return null;
  const rs=[];
  if(hitB) rs.push(`近8周合计${sum}台 <${T}台`);
  if(hitA) rs.push(`连续${maxrun}周 <${T}台/周`);
  return { reason: rs.join('；'), ex:f5Excp(m), T };
}
function f5Spark(m) {
  const w = (m.gro||{}).w8 || [];
  if (!w.length) return '<span class="na">待补</span>';
  const cl = f5Clear(m);
  const flat = w.every(v=>!v);   // 近8周全为0 → 渲染扁平灰条，保持柱状图形态
  const priceTxt = m.avg_price ? ` · 均价 ¥${m.avg_price}` : '';
  const bars = w.map((v,idx) => {
    if (v<=0) return `<i class="z" style="height:0;min-height:0;background:transparent" title="26W${29+idx}: 无销量"></i>`;  // 0销量周 → 留空位，不画柱
    const h = Math.max(2, Math.round((v/_w8max)*80));
    return `<i title="26W${29+idx}: ${v}台${priceTxt}" style="height:${h}px"${flat?' class="z"':''}></i>`;
  }).join('');
  const clTxt = cl ? `<em class="f5clr" title="清库：${cl.reason}${cl.ex?`（高价位例外，阈值${cl.T}台）`:''}">清库</em>` : '';
  return `${clTxt}<span class="f5spark${flat?' flat':''}"${flat?' title="近8周无销量"':''}>${bars}</span>`;
}
// 机会判定 + 对标表
function f5Draw() {
  const list = F5POOL.filter(m => F5.b.has(m.brand) && F5.cells.has(F5SZ_OF(m)+'||'+F5PR_OF(m)) && isClearShow(5, m));
  const n = list.length;
  const vol = list.reduce((s,m)=>s+m.cum_vol,0);
  // 近8周趋势：格内机型逐周销量汇总 26W30-37，后4周(34-37)相对前4周(30-33)
  const t8 = [0,0,0,0,0,0,0,0];
  list.forEach(m => { const w=(m.gro||{}).w8||[]; w.forEach((v,i)=>{ t8[i]+=(v||0); }); });
  const fW = t8.slice(0,4).reduce((a,b)=>a+b,0), lW = t8.slice(4).reduce((a,b)=>a+b,0);
  const trendR = (fW+lW) > 0 ? (fW>0 ? (lW-fW)/fW : null) : null;
  // 选中格子集合 → 展示标签（可多尺寸）
  const sd = new Set(), pd = new Set();
  F5.cells.forEach(c=>{ const sp=c.split('||'); sd.add(sp[0]); pd.add(sp[1]); });
  const sizeLb = [...sd].join('/'), prLabel = [...pd].join(' / ');
  const locLb = `${sizeLb||'?'}吋 × ${prLabel||'?'}`;
  const buckets = f5Buckets();
  let cls, tt, ds;
  const growing = trendR != null && trendR >= 0.20 && !(n >= buckets.dense);
  const crowded = n >= buckets.dense;
  const cfgTxt = (()=>{ if(!list.length) return '标配';
    const ss=list.filter(m=>!f5NoScore(m)).map(f5Score).filter(v=>!isNaN(v)).sort((a,b)=>a-b);
    return f5TierOf(ss[Math.floor(ss.length/2)], f5Tiers(list)).label; })();
  if (n === 0) { cls='gap'; tt='找空白（空档）';
    ds=`该「${locLb}」位置${'目前没有在售热门机型'}。属未被占位机会，但也说明需求可能未验证。建议看相邻价位/尺寸的体量，判断是否值得布局。`; }
  else if (crowded) { cls='opp'; tt='扎堆卷（红海）';
    ds=`该位置现有 ${n} 款、累计 ${fmt(vol)} 台，已饱和。你选「${cfgTxt}」切入，＋需靠价格或参数越级才能挤进前部；若无成本优势慎入。`; }
  else if (growing) { cls='grow'; tt='看增长（顺势）';
    ds=`该位置${n}款、累计 ${fmt(vol)} 台，近8周销量${trendR!=null?('环比 '+f5pct(trendR)):'待补'}（后4周 vs 前4周）。有量在涌入，以「${cfgTxt}」配置更快跟进，可吃到增量。`; }
  else { cls='mid'; tt='中等竞争（可切入）';
    ds=`该位置 ${n} 款、不拥挤，有正常在售。可按「${cfgTxt}」定位对标格内最强的 ${shortName(list[0])} 打差异化（价位/配置其一占优即可）。`; }
  const badge = document.getElementById('f5badge');
  badge.className = 'f5badge '+cls;
  badge.innerHTML = `<div class="bt">机会判定 · ${tt}</div>${ds}`;
  // 对标表
  document.getElementById('f5tcap').textContent = F5.cells.size>1 ? `对标表 · 多选 ${F5.cells.size} 处 (${locLb})` : `对标表 · ${locLb}`;
  const tbl = document.getElementById('f5tbl');
  // 累计销量动态周数
  const WE = D.meta ? (D.meta.week_end||35) : 35;
  const cols = [
    {k:'brand',t:'品牌'},{k:'model',t:'型号/产品名',f:m=>shortName(m)},{k:'size',t:'尺寸'},
    {k:'cum_vol',t:`26W1-W${WE}累计销量`,fmt:m=>fmtN(m.cum_vol)},{k:'avg_price',t:'均价',fmt:m=>'¥'+m.avg_price},
    {k:'score',t:'配置档',flt:true},
    {k:'trend',t:`近8周趋势`,f:m=>f5Spark(m)},
    {k:'tech',t:'屏幕技术',flt:true},{k:'part',t:'分区',flt:true},{k:'refresh',t:'原生刷新率',flt:true},
    {k:'anti',t:'抗反射',flt:true},{k:'mem',t:'内存',flt:true},{k:'audio',t:'音响',flt:true},
  ];
  // 列筛选应用
  let rows = list.slice();
  cols.filter(c=>c.flt).forEach(c => {
    const se = F5.filter[c.k];
    if (c.k==='score' && se && se.size===3) return;   // 配置档三档全选=不过滤 → 保留配置待补(f5NoScore)型号，避免对标表空档
    if (se && se.size) rows = rows.filter(m => se.has(f5Fval(c.k, m)));
  });
  if (F5.sort) {
    const sk = F5.sort;
    if (sk==='score') {
      // 配置档列排序：主按分值（升/降依表头），平分按销量降序
      rows.sort((a,b) => {
        const xs=f5Score(a), ys=f5Score(b);
        if (xs!==ys) return (F5.asc?1:-1)*(xs<ys?-1:1);
        return (b.cum_vol??0)-(a.cum_vol??0);
      });
    } else {
      rows.sort((a,b) => {
        let x,y;
        if (sk==='brand') { x=a.brand; y=b.brand; }
        else if (sk==='cum_vol') { x=a.cum_vol??0; y=b.cum_vol??0; }
        else if (sk==='avg_price') { x=a.avg_price??0; y=b.avg_price??0; }
        else if (sk==='model') { x=shortName(a); y=shortName(b); }
        else if (sk==='trend') { x=((a.gro||{}).w8||[]).reduce((s,v)=>s+v,0); y=((b.gro||{}).w8||[]).reduce((s,v)=>s+v,0); }
        else { x=f5Fval(sk,a); y=f5Fval(sk,b); if (sk==='refresh'||sk==='mem'){x=parseFloat(x)||0;y=parseFloat(y)||0;} }
        return (F5.asc ? 1 : -1) * (x<y?-1:x>y?1:0);
      });
    }
  } else if (F5.filter['score'] && F5.filter['score'].size===1) {
    // 仅筛选单一配置档：按分值降序，平分按销量降序
    rows.sort((a,b) => {
      const xs=f5Score(a), ys=f5Score(b);
      if (xs!==ys) return ys-xs;
      return (b.cum_vol??0)-(a.cum_vol??0);
    });
  }
  const tiers = f5Tiers(list); F5._tiers = tiers;
  // 左栏三档按钮：标注格内各档数量，选中态与对标表「配置档」筛选(F5.filter['score'])联动
  if (!F5.filter['score'] || F5.filter['score'].size === 0) F5.filter['score'] = new Set(['高配','标配','低价']);
  {
    const stg = {高配:0, 标配:0, 低价:0};
    list.forEach(m => { if(f5NoScore(m)) return; const L=f5TierOf(f5Score(m),tiers).label; if (stg[L]!=null) stg[L]++; });
    const cfgEl = document.getElementById('f5cfg');
    const ff = F5.filter['score'];
    if (cfgEl) cfgEl.querySelectorAll('span').forEach(sp => {
      const lb = ({'1':'高配','0':'标配','-1':'低价'})[sp.dataset.v];
      sp.innerHTML = `${lb} ${stg[lb]||0}`;
      sp.classList.toggle('on', ff.has(lb));
    });
  }
  let h = `<thead><tr>${cols.map(c=>{
    const fset = F5.filter[c.k]; const hasF = c.flt && fset && fset.size>0;
    return `<th data-s5="${c.k}">${c.flt?`<button class="f5fbtn${hasF?' act':''}" data-fk="${c.k}" title="筛选">${hasF?('▾·'+fset.size):'▾'}</button>`:''}<span>${c.t}${F5.sort===c.k?(F5.asc?' ▲':' ▼'):''}</span></th>`;
  }).join('')}</tr></thead><tbody>`;
  if (!rows.length) h += `<tr><td colspan="${cols.length}" class="na" style="text-align:center;padding:16px">无在售热门型号 —— 空档位置</td></tr>`;
  else rows.forEach(m => {
    const c = CO[m.brand]||'#0E7CE8';
    h += `<tr>${cols.map(c2=>{
      if (c2.k==='brand') return `<td style="border-left:3px solid ${c}">${m.brand}</td>`;
      if (c2.k==='size') return `<td>${m.size}吋</td>`;
      if (c2.k==='cum_vol') return `<td>${fmtN(m.cum_vol)}</td>`;
      if (c2.k==='trend') return `<td>${f5Spark(m)}</td>`;
      if (c2.k==='avg_price') return `<td>¥${m.avg_price}</td>`;
      if (c2.k==='score') { if(f5NoScore(m)) return `<td><span class="na">待补</span></td>`; const sc=f5Score(m); const tg=f5TierOf(sc,tiers); return `<td><span class="f5tier ${tg.cls}" title="配置分 ${sc}/100">${tg.label} ${sc}</span></td>`; }
      if (c2.k==='part') { const pv = m.part!=null && m.part!=='' ? m.part : (m.part_band ? (m.part_band==='无分区' ? '无' : m.part_band) : ''); return `<td>${pv ? pv : '<span class="na">待补</span>'}</td>`; }
      if (c2.k==='refresh') return `<td>${f5Hz(m.refresh_hz)}</td>`;
      if (c2.k==='mem') return `<td>${memFmt(m.mem)}</td>`;
      if (c2.k==='tech') return `<td>${f5Tech(m)}</td>`;
      if (c2.k==='anti') { const a=m.anti||m.anti_band; return `<td>${a||'<span class="na">待补</span>'}</td>`; }
      if (c2.k==='audio') return `<td>${m.audio||'<span class="na">待补</span>'}</td>`;
      if (c2.k==='model') {
        const tip = m.product_name ? String(m.model).replace(/"/g,'&quot;') : '';
        return `<td class="f5anchor" title="${tip}">${shortName(m)}${gDot(m)}${isNew(m)?'<span class="ngchip">新品</span>':''}</td>`;
      }
      return `<td>${m[c2.k]==null?'<span class="na">待补</span>':m[c2.k]}</td>`;
    }).join('')}</tr>`;
  });
  h += `</tbody>`;
  tbl.innerHTML = h;
  buildFloatHeads();   // 对标表表头吸顶（复用四图浮动表头机制）
  const dlBtn = document.getElementById('f5dlBtn');
  if (dlBtn) dlBtn.onclick = () => exportF5Csv(list);
  f5Deep(list, locLb);
  tbl.querySelectorAll('th[data-s5]').forEach(th => th.onclick = (e) => {
    if (e.target.closest('.f5fbtn')) return;
    const k = th.dataset.s5;
    if (F5.sort === k) F5.asc = !F5.asc; else { F5.sort = k; F5.asc = true; }
    f5Draw();
  });
  /* 列筛选：点击 ▾ 打开该列的下拉多选面板 */
  document.querySelectorAll('.f5fbtn').forEach(btn => btn.onclick = (e) => {
    e.stopPropagation();
    f5OpenDD(btn.dataset.fk, btn, list);
  });
  /* 双击明细行 → 加入对比模式 */
  tbl.querySelectorAll('tbody tr').forEach(tr => tr.ondblclick = () => {
    const modelName = tr.cells[1] ? tr.cells[1].textContent : '';
    const m = list.find(x => shortName(x) === modelName);
    if (!m || S.sel.length >= 8) return;
    S.sel.push(m.id);
    S.cmp = true; document.getElementById('cmpBtn').classList.add('on');
    document.getElementById('cmpBar').classList.add('show');
    document.body.classList.add('cmpOn');
    renderCmpBar();
  });
}
function f5CloseDD() {
  const d = document.getElementById('f5ddown');
  if (d) d.classList.remove('show');
}
function f5OpenDD(k, btn, list) {
  const d = document.getElementById('f5ddown');
  const mp = new Map();
  list.forEach(m => { const v = f5Fval(k, m); mp.set(v, (mp.get(v)||0)+1); });
  const vals = [...mp.keys()].sort();
  const cur = F5.filter[k] || new Set();
  const items = vals.map(v =>
    `<label class="f5dd-it ${cur.has(v)?'on':''}"><input type="checkbox" value="${v}" ${cur.has(v)?'checked':''}><span>${v}</span><em>${mp.get(v)}</em></label>`).join('');
  d.innerHTML = `<div class="f5dd-t"><span>${k}</span><span class="f5dd-x" data-x title="关闭">✕</span></div>
    <div class="f5dd-list">${items||'<div style="padding:6px;color:#94a3b8;font-size:11px">无选项</div>'}</div>
    <div class="f5dd-f"><button data-dd="all">全选</button><button data-dd="none">清空</button><button data-dd="done">完成</button></div>`;
  const r = btn.getBoundingClientRect();
  d.style.left = Math.max(4, Math.min(r.left, window.innerWidth-220)) + 'px';
  d.style.top = Math.max(4, r.bottom + 4) + 'px';
  d.classList.add('show');
  d._k = k;
  d.querySelector('.f5dd-list').onchange = (e) => {
    const inp = e.target;
    if (inp.tagName !== 'INPUT') return;
    F5.filter[k] = F5.filter[k] || new Set();
    const s = F5.filter[k];
    if (inp.checked) s.add(inp.value); else s.delete(inp.value);
    if (!s.size) delete F5.filter[k];
    inp.closest('label').classList.toggle('on', inp.checked);
    f5Draw();   // 重绘表格（下拉面板独立于表格，保持打开）
  };
  d.querySelector('[data-x]').onclick = f5CloseDD;
  d.querySelector('[data-dd="all"]').onclick = () => { F5.filter[k] = new Set(vals); f5CloseDD(); f5Draw(); };
  d.querySelector('[data-dd="none"]').onclick = () => { delete F5.filter[k]; f5CloseDD(); f5Draw(); };
  d.querySelector('[data-dd="done"]').onclick = f5CloseDD;
}
function buildFig5() {
  f5TargetChips();
  f5BrandFilter();
  buildClearCtl(5, 'fig5brand');
  f5Grid();
  f5Draw();
  // 点击面板外 / Esc 关闭下拉筛选
  document.addEventListener('click', e => {
    const d = document.getElementById('f5ddown');
    if (!d || !d.classList.contains('show')) return;
    if (e.target.closest('#f5ddown') || e.target.closest('.f5fbtn')) return;
    f5CloseDD();
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') f5CloseDD(); });
}
buildFig5();
(function(){var b=document.getElementById('toTop');if(!b)return;var t=function(){b.style.display=(window.scrollY>600)?'flex':'none';};window.addEventListener('scroll',t,{passive:true});t();b.onclick=function(){window.scrollTo({top:0,behavior:'smooth'});};})();
"""

# ============ HTML ============
def build_html():
    payload = json.dumps(DATA, ensure_ascii=False, separators=(",", ":"))
    n = len(models)
    spec_n = sum(1 for m in models if m.get("spec_src") == "ref")
    ser_n = len({m["series"] for m in models})
    # 图1 布局口径：近20周(26W18-W37) 主流(尺寸<98)>1000 / 98吋+(≥98)>200 台 且 尺寸能落进图1 的 9 个尺寸列
    g1 = [m for m in models if m.get("fig1_ok")]
    n1 = len(g1)
    ser1_n = len({m["series"] for m in g1})
    total1_v = sum(m["cum_vol"] for m in g1)

    def _gate(m):
        return m["cum_vol"] > 5000 if m["size"] < 98 else m["cum_vol"] > 1000
    g3 = [m for m in models if m["size"] > 50 and _gate(m)]
    g4 = [m for m in models if m["size"] <= 50 and m["cum_vol"] > 1000]
    param_n = len(g3) + len(g4)
    total_v = sum(m["cum_vol"] for m in models)
    # 图4 维度覆盖（真实口径，替代 v3 时代「能效全缺」的过时文案）
    e_cov = sum(1 for m in g4 if m.get("energy"))
    m_cov = sum(1 for m in g4 if m.get("mem_band"))
    r_cov = sum(1 for m in g4 if m.get("res_band"))
    # 图3 参数覆盖
    p_cov = sum(1 for m in g3 if m.get("part_band"))
    hz_cov = sum(1 for m in g3 if m.get("refresh_hz"))
    # footer 动态数字
    n60 = sum(1 for m in models if m["size"] == 60)
    mi_f1 = sum(1 for m in g1 if m["brand"] in ("小米", "红米"))
    today = datetime.date.today().isoformat()

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>8品牌电视 AVC 定位布局报告</title>
<style>{CSS}</style>
</head>
<body>
<div class="wrap">

<header>
  <h1>8品牌电视定位布局报告</h1>
  <div class="sub">
    图1 布局口径：<b>AVC {META['fig1_window']} 主流&gt;1000 / 98吋+&gt;200 台</b>（{n1} 款 · {ser1_n} 个系列）　
    销量口径：<b>AVC 26W01–26W{META['week_end']} 累计</b>（全量 {n} 款 · 累计 {total_v/10000:.1f} 万台）　
    图3／图4 门槛（近20周）：图3 &gt;50吋 主流&gt;3000 / 98吋+&gt;500 台；图4 ≤50吋 &gt;1000 台（共 {param_n} 款）
  </div>
</header>

<div class="ctrl">
  <div class="crow">
    <span class="clab">筛选</span>
    <input class="search" id="q" placeholder="搜索型号 / 系列 / 品牌 / 拼音，如 P5H、鹤6、hx=海信">
    <span class="btn" id="csvBtn">⬇ 导出 CSV</span>
    <span class="btn" id="cmpBtn">⚖ 对比模式</span>
    <span class="sum">当前 <b id="statN">0</b> 款 · 累计 <b id="statV">0</b> 台</span>
  </div>
</div>

<!-- 新品上市板块（近4周滚动首次出现的型号） -->
<div class="newsec" id="newsec" style="display:none">
  <div class="hd">
    <h2>新品上市</h2>
    <span class="cnt" id="newcnt"></span>
  </div>
  <div class="sub" id="newsub"></div>
  <div class="newgrid" id="newgrid"></div>
</div>

<section>
  <h2>图1　系列 × 尺寸 销量分布</h2>
  <div class="fig-brand" id="fig1brand"></div>
  <div class="note">9 个尺寸（<b>60吋并入 55、100吋+ 并入 98吋+</b>）
    入选门槛 <b>近20周(26W18–W37) 主流&gt;1000 / 98吋+&gt;200 台</b>　底色深浅 = 该格销量强度，入选 {n1} 款 / {ser1_n} 系列。</div>
  <div class="scroll" id="fig1body"></div>
  <div class="legend">
    <span class="lg"><span class="heat-scale"><span class="hs-bar"></span>浅 → 深 = 单格销量强度（0 → 峰值）</span></span>
    <span id="fig1cnt"></span>
  </div>
</section>

<div id="expbox"></div>

<section>
  <h2>图2　均价价格带 × 尺寸 分布 <span class="tag">点击格子展开型号</span>
    <span class="btn mini" id="fig2exp">⊕ 展开全部</span></h2>
  <div class="fig-brand" id="fig2brand"></div>
  <div class="note">行 = 12 档均价带（左闭右开），列 = 尺寸（<b>98-100吋与100吋以上合并为「98吋+」</b>）。
    格内 = 累计销量与型号数，底部色条 = <b>品牌份额（TOP3 + 其他）</b>，悬停看全部品牌明细。
    点击格子展开型号，面板顶部<b>品牌 chips 可筛选</b>（多选，只作用于本格）。</div>
  <div class="scroll" id="fig2body"></div>
  <div class="legend"><span id="fig2cnt"></span></div>
</section>

<section>
  <h2>图3　主力型号配置图（&gt;50吋）</h2>
  <div class="fig-brand" id="fig3brand"></div>
  <div class="fig-brand" id="fig3size"></div>
  <div class="legend mem-legend">
    <span class="lg"><i style="background:#e1f5fe;border:1px solid #dde2e8"></i>内存 ≤2GB</span>
    <span class="lg"><i style="background:#b3e5fc;border:1px solid #dde2e8"></i>3GB</span>
    <span class="lg"><i style="background:#81d4fa;border:1px solid #dde2e8"></i>4GB+</span>
    <span class="lg"><i style="background:#f6f7f9;border:1px dashed #c9cfd8"></i>待补</span>
  </div>
  <div class="note">行 = 控光分区档（高档按分区数细分：<b>三千级／四千-七千级／八千级+</b>），列 = <b>刷新率</b>（120Hz 与 132Hz 合并为 120-132Hz；165Hz 与 170Hz 合并为 165-170Hz；余 60／144／150／180Hz）。
    芯片左色条 = 品牌色，芯片底色 = 内存档，徽章 = 抗反射（AG／LR）+ 屏幕技术（Mini LED／QD-MiniLED／SQD-MiniLED／RGB-MiniLED／QLED 等进阶技术显示原名并高亮，普通液晶统一示作 LED）。门槛（近20周）：&gt;50吋 主流&gt;3000台、98吋+&gt;500台（{len(g3)} 款；分区已覆盖 {p_cov} 款、刷新率已覆盖 {hz_cov} 款，其余「待补」占位）。</div>
  <div id="fig3tools" class="ftools"></div>
  <div class="scroll" id="fig3body"></div>
  <div class="legend"><span id="fig3cnt"></span></div>
</section>

<section>
  <h2>图4　小尺寸配置图（≤50吋）</h2>
  <div class="fig-brand" id="fig4brand"></div>
  <div class="fig-brand" id="fig4size"></div>
  <div class="note">行 = 能效（一／二级）或内存档，列 = 分辨率（4K／1080P／WXGA）。门槛：≤50吋且 &gt; 1000 台（{len(g4)} 款）。
    当前覆盖：能效 {e_cov} 款、内存 {m_cov} 款、分辨率 {r_cov} 款，其余以「待补」占位。</div>
  <div class="warn" id="fig4warn" style="display:none"></div>
  <div id="fig4tools" class="ftools"></div>
  <div class="scroll" id="fig4body"></div>
  <div class="legend"><span id="fig4cnt"></span></div>
</section>

<section>
  <h2>图5　产品定位分析</h2>
  <div class="fig-brand" id="fig5brand"></div>
  <div class="note">以「尺寸 × 价位段」为竞争地图，定位<b>你的某一款产品</b>落点（仅分析 <b>&gt;50 吋</b>大屏市场），看它吃哪种市场机会（<b>扎堆卷 / 看增长 / 找空白</b>）再下钻对标。
    格色 = 该位置拥挤度（过门槛型号数，自动分桶），蓝标 = 该格近8周销量上升（后4周 vs 前4周增长≥20%），红/深色格 = 已选中的目标落点（<b>点击可多选，再点取消</b>）。门槛（近20周）：<b>98吋+ &gt;200 台、其他尺寸 &gt;1000 台</b>。趋势口径：AVC 26W30–37 逐周销量。</div>
  <div class="f5wrap">
    <div class="f5left">
      <div class="f5card">
        <h4>① 目标产品</h4>
        <div class="f5field"><label>尺寸档</label><div class="f5seg" id="f5sz"></div></div>
        <div class="f5field"><label>价位段（元）</label><div class="f5seg" id="f5pr"></div></div>
        <div class="f5field"><label>配置定位（对同片对手）</label>
          <div class="f5seg" id="f5cfg"><span data-v="1">高配</span><span data-v="0">标配</span><span data-v="-1">低价</span></div>
        </div>
      </div>
      <div class="f5badge" id="f5badge"><div class="bt">机会判定</div>点选左侧尺寸 / 价位段，查看落点机会与打法。</div>
      <div class="f5deep" id="f5deep"></div>
    </div>
    <div class="f5main">
      <div class="f5maptitle">竞争密度地图 · 横轴=价位段（&lt;1999 → 16000+）　纵轴=尺寸（点击格子定位，可多点价位段合并对标）</div>
      <div class="f5map" id="f5map"></div>
      <div class="f5legend">
        <span class="lg"><i class="f5sw" style="background:#f1f5f9"></i>空白</span>
        <span class="lg"><i class="f5sw" style="background:#dbeafb"></i>稀疏</span>
        <span class="lg"><i class="f5sw" style="background:#93c5fd"></i>中等</span>
        <span class="lg"><i class="f5sw" style="background:#1d4ed8"></i>密集</span>
        <span class="lg"><i style="width:16px;height:16px;line-height:16px;text-align:center;font-size:9px;color:#fff;background:#2563eb;display:inline-block;border-radius:3px">▲</i>增长</span>
        <span class="lg"><i style="width:12px;height:12px;border-radius:2px;border:2px solid #ef4444;background:transparent;display:inline-block"></i>你的产品</span>
      </div>
      <div class="f5topline" style="margin-top:10px">
        <span class="tcap" id="f5tcap">对标表</span>
        <span style="flex:1"></span>
        <button class="btn mini" id="f5dlBtn" title="下载竞争地图当前选中格内的型号参数（含品牌/尺寸筛选，不含列筛选）">⬇ 下载选中型号</button>
      </div>
      <div class="f5ddown" id="f5ddown"></div>
      <div class="scroll"><table class="f5tbl" id="f5tbl"></table></div>
      <div class="f5score">
        <b>配置评分规则</b>（总分100 = 屏幕技术35 + 分区25 + 刷新率10 + 内存10 + 抗反射10 + 音响10）
        <div>
          <span>技术35：OLED 35 / RGB-Mini LED 30 / SQD-Mini LED 28 / BGB-Mini LED 20 / QD-Mini LED 18 / Mini LED 15 / QLED 5 / LED 0</span>
          <span>分区25：无0 → 1-99 3 → 100-199 5 → 200-499 7 → 500-699 10 → 700-999 12 → 1000-1499 16 → 1500-1999 18 → 2000-2999 19 → 3000-3999 20 → 4000-7999 22 → 8000+ 25</span>
          <span>刷新率10：60Hz 2 / 120Hz 4 / 132–144Hz 6 / 150Hz 7 / 165Hz 8 / 170Hz 9 / 180Hz 10</span>
          <span>内存10：RAM 1=0 / 1.5=2 / 2=4 / 3=5 / ≥4=6 ＋ ROM 8=0 / 16=1 / 32=2 / 64=3 / ≥128=4（1+8=0…4+128=10）</span>
          <span>抗反射10：无0 / AG 4 / LR+AG 8 / LR 10</span>
          <span>音响10：2.0声道0 / 2.1=4 / 2.1.2=6 / 更高=8 / 更高且名品(安桥·帝瓦雷·哈曼…)10</span>
          <span>档位口径＝该格子内配置分经 K-Means 聚 3 簇，按最接近簇判定高配/标配/低价</span>
        </div>
      </div>
    </div>
  </div>
</section>

<footer>
  数据来源：AVC 26W01–W{META['week_end']} 明细（入选门槛近20周 {META['fig1_window']}）· 系列划分与型号参数取自《海信TCL产品布局分析》　|　
  生成于 {today}　|　60吋（{n60} 款）并入 55吋列；27吋（非电视尺寸档）已在数据阶段剔除；小米／红米（图1 入选 {mi_f1} 款）参数待补，补齐后自动生效
</footer>
</div>

<div id="cmpBar">
  <div class="cb-in">
    <span class="cb-t">型号对比 <span id="cmpN">0</span>/8</span>
    <span class="cb-list" id="cmpList"></span>
    <button class="btn" id="goCmp" disabled>开始对比</button>
    <span class="btn" id="cmpClear">清空</span>
    <span class="btn" id="cmpExit">退出对比</span>
  </div>
</div>

<div class="modal" id="cmpModal">
  <div class="mbox">
    <div class="mhd">
      <h3>型号参数对比</h3>
      <label style="font-size:12px;color:#6b7280"><input type="checkbox" id="diffOnly"> 仅显示差异</label>
      <span class="x" id="mClose">✕</span>
    </div>
    <div class="mbd" id="cmpBody"></div>
  </div>
</div>

<script>window.__DATA__ = {payload};</script>
<button id="toTop" title="回到顶部" style="position:fixed;right:22px;bottom:88px;width:46px;height:46px;border-radius:50%;border:1px solid #d8dde5;background:#fff;color:#1f2328;font-size:22px;line-height:1;box-shadow:0 4px 14px rgba(15,23,42,.18);cursor:pointer;z-index:9000;padding:0;display:flex;align-items:center;justify-content:center;">↑</button>
<script>{JS}</script>
</body>
</html>"""
    return html

if __name__ == "__main__":
    html = build_html()
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"已生成 {OUT}  ({len(html)/1024:.1f} KB)")
