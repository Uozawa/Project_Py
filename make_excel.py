#!/usr/bin/env python3
"""塗料メーカー 製品タイプ別 製造フロー IN/OUT 表 Excel 生成スクリプト

添付 PDF「リン酸(湿式法) IN/OUT表(マテリアルバランス)」のテンプレート構造を踏襲。
レイアウト:
  上段 = 各工程の IN ブロック (品名 / 標準成分 / 量 / 備考)
  中段 = 工程名と概要
  下段 = 各工程の OUT ブロック (同上)
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
)
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "塗料メーカー_製造フローINOUT.xlsx"

# ────────────────────────────────────────────────────────────────────
# スタイル
# ────────────────────────────────────────────────────────────────────
thin = Side(style="thin", color="666666")
medium = Side(style="medium", color="333333")
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
BORDER_BOX = Border(left=medium, right=medium, top=medium, bottom=medium)

FILL_TITLE = PatternFill("solid", fgColor="1F3864")
FILL_IN = PatternFill("solid", fgColor="DDEBF7")          # 薄青
FILL_OUT = PatternFill("solid", fgColor="FCE4D6")         # 薄橙
FILL_PROCESS = PatternFill("solid", fgColor="E2EFDA")     # 薄緑
FILL_HEADER_IN = PatternFill("solid", fgColor="9DC3E6")
FILL_HEADER_OUT = PatternFill("solid", fgColor="F4B084")
FILL_HEADER_PROC = PatternFill("solid", fgColor="A9D08E")

FONT_TITLE = Font(name="Yu Gothic", size=14, bold=True, color="FFFFFF")
FONT_H = Font(name="Yu Gothic", size=10, bold=True)
FONT_BODY = Font(name="Yu Gothic", size=9)
FONT_PROC = Font(name="Yu Gothic", size=10, bold=True)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

# ────────────────────────────────────────────────────────────────────
# データ定義
# ────────────────────────────────────────────────────────────────────

PRODUCT_TYPES = [
    {
        "id": "①",
        "name": "溶剤系液体塗料",
        "用途代表例": "自動車補修・工業用・建築鉄部・防食",
        "特徴": "有機溶剤を媒体とする。乾燥が早く塗膜性能が高いがVOC排出大",
        "規制": "消防法 第4類引火性液体 / 大気汚染防止法 (VOC)",
    },
    {
        "id": "②",
        "name": "水性液体塗料",
        "用途代表例": "建築用 (外壁・内装)・家庭用DIY・水性自動車塗料",
        "特徴": "水を媒体とする。低VOC。凍結NG・防腐剤必要",
        "規制": "VOC削減型として推奨。冬季の保管温度管理が必要",
    },
    {
        "id": "③",
        "name": "粉体塗料",
        "用途代表例": "家電筐体・自動車部品・建材・パイプ",
        "特徴": "溶剤を全く使わない粉末。回収再利用可能・塗着効率ほぼ100%",
        "規制": "VOCゼロ。ただし押出機・粉砕機の設備投資大",
    },
    {
        "id": "④",
        "name": "電着塗料",
        "用途代表例": "自動車車体下塗 (カチオン)・電子部品 (アニオン)",
        "特徴": "水中で電気泳動により塗装。下塗の世界標準",
        "規制": "塗装ラインと一体運用。塗料は水性分散体として供給",
    },
    {
        "id": "⑤",
        "name": "UV硬化塗料",
        "用途代表例": "木工 (フローリング・家具)・プラスチック・印刷",
        "特徴": "紫外線照射で瞬時に硬化。溶剤レス・省エネ",
        "規制": "光開始剤含有のため遮光保管。製造ラインも遮光下",
    },
]


# 各製品タイプの工程ステップと IN / OUT
# 各 step: {name, desc, in_items: [(品名,成分,量,備考), ...], out_items: [...]}
# 量はバッチサイズ 10,000 kg を想定した代表値 (概算)。出典: 業界一般値 + 推定。

FLOWS = {

# ─────────── ① 溶剤系液体塗料 ───────────
"溶剤系液体塗料": [
    {
        "name": "①計量・配合 (Pre-weighing / Charging)",
        "desc": "処方 (Recipe) に従って樹脂・顔料・溶剤・添加剤をロードセル付き計量タンクで秤量し、釜に投入。粉体顔料は自動ホッパーで搬送、液体はポンプ計量。所要時間 30〜60 分/バッチ。",
        "in": [
            ("樹脂 (アクリル/アルキッド/エポキシ等)", "固形分 40〜60%", "4,200 kg", "塗膜骨格を形成 (バインダー)"),
            ("顔料 (TiO₂・有機顔料)", "TiO₂ 等", "1,800 kg", "着色・隠蔽・機能付与"),
            ("有機溶剤 (トルエン/キシレン/MEK 等)", "引火点 21〜70℃ (第1〜2石油類)", "3,500 kg", "粘度調整・乾燥制御"),
            ("添加剤 (分散剤/消泡剤/レベリング剤)", "BYK等", "500 kg", "品質向上・少量添加"),
            ("電力 (秤・搬送機)", "-", "微", "設備動力"),
        ],
        "out": [
            ("配合済バッチ (釜内)", "樹脂+顔料+溶剤+添加剤の集合", "10,000 kg", "次工程へ"),
            ("空袋・空ドラム", "PE/PP/紙袋・金属ドラム", "数十 kg", "汚染あれば産廃、清浄なら金属スクラップ"),
            ("粉塵 (顔料投入時)", "TiO₂等微粒子", "微量", "集塵機で回収→産廃"),
            ("計量誤差バッチ", "-", "0〜微量", "誤差時は再計量で復元"),
        ],
    },
    {
        "name": "②予備混合 (Pre-mixing)",
        "desc": "ディスパーザー (高速回転羽根 1,500〜3,000 rpm) で粗分散。所要時間 30〜60 分。窒素シールで爆発防止。",
        "in": [
            ("配合済バッチ", "前工程産物", "10,000 kg", ""),
            ("電力 (高速攪拌 5〜30 kW)", "Scope 2", "5〜30 kWh", "ディスパー駆動"),
            ("冷却水 (ジャケット)", "産業用水", "0.5 m³/t", "発熱抑制"),
            ("窒素ガス (シール)", "N₂", "微量", "防爆"),
        ],
        "out": [
            ("粗分散ペースト", "ミルベース前駆体", "9,990 kg", "次工程へ"),
            ("反応熱・摩擦熱", "顕熱", "-", "ジャケット排水で除去"),
            ("温水排水", "工業排水", "-", "クーリングタワー循環"),
            ("VOC 蒸気", "溶剤揮散", "微量", "RTO/活性炭で処理"),
            ("騒音・振動", "-", "-", "労働環境課題"),
        ],
    },
    {
        "name": "③分散・練肉 (Dispersion / Milling)",
        "desc": "サンドミル/ビーズミル/三本ロールミルで顔料を nm 単位まで微粒化。最も電力を消費する工程。所要時間 1〜数時間。",
        "in": [
            ("粗分散ペースト", "前工程産物", "9,990 kg", ""),
            ("ビーズメディア (補充分)", "ジルコニア・ガラス 0.3〜2 mm", "数 kg", "摩耗分の補充"),
            ("電力 (10〜100 kW)", "Scope 2 大消費", "業界一般値の概算", "粒径低下のためのエネルギー"),
            ("冷却水", "産業用水", "2〜5 m³/t (概算)", "発熱が大きい工程"),
        ],
        "out": [
            ("ミルベース (微分散塗料)", "粒径 数十〜数百 nm", "9,970 kg", "次工程へ"),
            ("摩耗ビーズ廃材", "ジルコニア等", "数 kg", "産廃"),
            ("大量の発熱", "顕熱", "-", "冷却水で除去"),
            ("温水排水", "-", "-", ""),
            ("微量 VOC", "溶剤揮散", "微量", "RTO 処理"),
        ],
    },
    {
        "name": "④レットダウン (Let-down)",
        "desc": "ミルベースに追加樹脂・希釈剤・添加剤を低速攪拌 (100〜300 rpm) で混合。塗料化工程。所要時間 30〜90 分。",
        "in": [
            ("ミルベース", "前工程産物", "9,970 kg", ""),
            ("追加樹脂・希釈剤", "アクリル等+ 溶剤", "→ 計算上含む", "粘度・固形分調整"),
            ("添加剤 (消泡・レベリング)", "後添加分", "→ 計算上含む", ""),
            ("低速攪拌動力", "Scope 2", "数 kWh", ""),
        ],
        "out": [
            ("配合完了塗料 (ベース)", "粘度・固形分が規格内", "9,960 kg", "次工程 調色へ"),
            ("撹拌熱", "顕熱", "微", ""),
            ("微量 VOC", "溶剤揮散", "微量", "RTO 処理"),
        ],
    },
    {
        "name": "⑤調色 (Tinting / Color Match)",
        "desc": "ベース塗料に着色剤 (Tinter) を分光測色計で計測しながら微調整。CCM (Computer Color Matching) を使用。色差 ΔE < 0.8〜1.5 で合格。",
        "in": [
            ("ベース塗料", "前工程産物", "9,960 kg", ""),
            ("着色剤 (Color Paste)", "高濃度有機/無機顔料分散体", "10〜200 kg", "色相微調整"),
            ("色見本・標準色板", "顧客指定色", "-", "目標値"),
            ("電力 (CCM/分光測色計)", "Scope 2", "微", ""),
        ],
        "out": [
            ("目標色の塗料", "ΔE < 0.8〜1.5", "9,990 kg", "次工程へ"),
            ("調色サンプル板", "試験塗装したテストピース", "数百 g/回", "QC 履歴として一定期間保管後 産廃"),
            ("サンプル廃材", "塗布残・洗浄液", "数百 g〜数 kg", "産廃"),
            ("補正バッチ (再調色)", "-", "発生時のみ", "コスト要因"),
        ],
    },
    {
        "name": "⑥ろ過 (Filtration)",
        "desc": "バッグフィルタ/カートリッジで凝集物・異物を除去。メッシュ目開き 50〜200 μm。圧力差で目詰まり監視。所要時間 10〜30 分。",
        "in": [
            ("調色完了塗料", "前工程産物", "9,990 kg", ""),
            ("フィルター", "ナイロン/PP 不織布", "数枚〜数十枚", "目詰まり時交換"),
            ("ポンプ動力", "Scope 2", "微", ""),
        ],
        "out": [
            ("清浄塗料", "凝集物除去済", "9,980 kg", "次工程へ"),
            ("使用済フィルター", "塗料含浸", "数 kg", "産廃 (汚染あり)"),
            ("凝集物・ゲル粒", "塗料カス", "数 kg", "産廃"),
        ],
    },
    {
        "name": "⑦品質検査 (QC Test)",
        "desc": "粘度 (フォードカップ/ストーマー)・比重・色差・隠蔽率・乾燥時間を測定。標準品 (Reference) と比較。所要 1〜3 時間 (乾燥試験は半日〜1日)。",
        "in": [
            ("塗料サンプル", "ろ過後塗料の一部", "0.5〜2 kg", ""),
            ("標準品 (Reference)", "前ロットまたは標準ロット", "微", ""),
            ("試薬・溶媒", "希釈用シンナー等", "微", ""),
            ("電力 (測定機器・乾燥炉)", "Scope 2", "微", ""),
        ],
        "out": [
            ("合格判定塗料 (本体)", "QC 通過", "9,980 kg", "充填工程へ"),
            ("QC データ・成績書", "ロット記録", "-", "出荷証憑"),
            ("サンプル廃液", "試験後残液", "0.5〜2 kg", "産廃"),
            ("不合格バッチ", "規格外", "発生時のみ", "補正 or 別グレード転用 or 廃棄"),
        ],
    },
    {
        "name": "⑧充填・包装 (Filling / Packaging)",
        "desc": "自動充填機 (重量式) で 1L / 4L / 14L / 16L / 20L (一斗缶) / 200L ドラム / IBC コンテナに充填。GHS 表示・ロット印字。所要時間 0.5〜数秒/缶。",
        "in": [
            ("合格塗料", "前工程産物", "9,980 kg", ""),
            ("空缶・キャップ・ラベル", "金属缶/プラ容器", "缶数 ≒ 製品量/容量", "包装資材"),
            ("段ボール・パレット", "外装", "-", ""),
            ("電力 (充填機・封缶機)", "Scope 2", "数 kWh", ""),
        ],
        "out": [
            ("製品缶", "充填完了製品", "9,950 kg", "出荷"),
            ("滴下塗料・残留", "充填損失", "10〜30 kg", "産廃"),
            ("不良缶・印字ミス", "再充填 or 廃棄", "少量", ""),
            ("包装廃材 (端材)", "段ボール等", "-", "紙リサイクル"),
            ("微量 VOC 揮散", "蓋閉前", "微量", ""),
        ],
    },
    {
        "name": "並行A: 設備洗浄 (Cleaning)",
        "desc": "色替え・釜替えごとに洗浄シンナーで内部洗浄。蒸留再生機 (Solvent Recovery) で廃シンナーをリサイクル。",
        "in": [
            ("洗浄シンナー", "回収再生品 or 新品", "100〜500 kg/回", "色替え毎"),
            ("電力・スチーム", "Scope 1・2", "数 kWh + 数十 kg", "温水洗浄"),
            ("人件費 (作業者)", "-", "-", "内部清掃"),
        ],
        "out": [
            ("廃シンナー (特管産廃)", "引火点 70℃未満の廃油", "100〜500 kg/回", "蒸留再生 or 焼却"),
            ("洗浄排水", "工業排水", "-", "排水処理プラント"),
            ("残塗料 (回収)", "次バッチ用 or 産廃", "数 kg〜数十 kg", ""),
            ("VOC 排出", "揮散溶剤", "工程内最大級", "RTO 処理"),
        ],
    },
],

# ─────────── ② 水性液体塗料 ───────────
"水性液体塗料": [
    {
        "name": "①計量・配合",
        "desc": "水性エマルション樹脂・顔料・水・添加剤を計量。ステンレス釜を使用 (錆対策)。",
        "in": [
            ("水性エマルション樹脂", "アクリル/ウレタン分散体 固形分 40〜55%", "4,500 kg", "水中分散体"),
            ("顔料 (TiO₂・有機顔料)", "TiO₂ 等", "2,000 kg", ""),
            ("水 (媒体)", "イオン交換水・純水", "2,500 kg", "VOC削減効果"),
            ("造膜助剤 (テキサノール等)", "微量溶剤", "100 kg", "低温造膜性付与"),
            ("添加剤 (分散剤/消泡/増粘/防腐剤)", "BYK・Lonza等", "400 kg", "★防腐剤必須"),
            ("pH 調整剤 (アンモニア/AMP)", "塩基", "微量", "pH 8〜9 維持"),
        ],
        "out": [
            ("配合済バッチ", "水分散体", "9,500 kg", "次工程へ"),
            ("空袋・空ドラム", "-", "数十 kg", ""),
            ("粉塵 (顔料投入時)", "TiO₂", "微量", "集塵機"),
        ],
    },
    {
        "name": "②予備混合・分散",
        "desc": "高速ディスパー → ビーズミル分散。発泡しやすいので消泡剤管理が重要。所要時間 1〜数時間。",
        "in": [
            ("配合済バッチ", "前工程産物", "9,500 kg", ""),
            ("電力 (高速分散)", "Scope 2", "業界一般値の概算", "顔料分散"),
            ("冷却水", "-", "-", "発熱抑制 (溶剤系より発熱小)"),
        ],
        "out": [
            ("ミルベース", "微分散水性塗料", "9,490 kg", "次工程へ"),
            ("摩耗ビーズ", "-", "微", "産廃"),
            ("温水排水", "-", "-", ""),
            ("水蒸気・微量 VOC", "造膜助剤が一部揮散", "微量", "排ガス処理"),
        ],
    },
    {
        "name": "③レットダウン・pH 調整",
        "desc": "追加エマルションを添加し、pH を 8〜9 に調整。防腐剤を最終添加。",
        "in": [
            ("ミルベース", "前工程産物", "9,490 kg", ""),
            ("追加エマルション・水", "-", "→ 量に含む", "粘度・固形分調整"),
            ("pH 調整剤・防腐剤", "アンモニア/AMP・イソチアゾリン系", "微量", ""),
            ("低速攪拌動力", "Scope 2", "微", ""),
        ],
        "out": [
            ("配合完了塗料 (ベース)", "pH 8〜9、粘度規格内", "9,500 kg", "調色へ"),
            ("微量アンモニア揮散", "-", "微量", "排気処理"),
        ],
    },
    {
        "name": "④調色",
        "desc": "水性着色剤で CCM 調色。乾燥色と湿潤色の差 (ドライダウン現象) を考慮。",
        "in": [
            ("ベース塗料", "前工程産物", "9,500 kg", ""),
            ("水性着色剤", "顔料分散体", "10〜200 kg", ""),
            ("CCM 装置", "-", "微電力", ""),
        ],
        "out": [
            ("目標色の水性塗料", "ΔE 規格内", "9,650 kg", ""),
            ("試験塗板・サンプル", "-", "数百 g", "産廃"),
        ],
    },
    {
        "name": "⑤ろ過",
        "desc": "細目フィルタで凝集物除去。エマルションの目詰まりに注意。",
        "in": [
            ("調色塗料", "前工程産物", "9,650 kg", ""),
            ("フィルター", "PP/PE 不織布", "数枚", ""),
        ],
        "out": [
            ("清浄塗料", "ろ過後", "9,640 kg", ""),
            ("使用済フィルター", "-", "数 kg", "産廃"),
        ],
    },
    {
        "name": "⑥品質検査",
        "desc": "粘度・比重・色差・隠蔽率に加え、凍結融解試験・耐水性試験など水性特有のチェック。",
        "in": [
            ("塗料サンプル", "-", "1〜2 kg", ""),
            ("試験設備", "電力", "微", ""),
        ],
        "out": [
            ("合格塗料", "-", "9,640 kg", ""),
            ("QC 記録", "-", "-", ""),
            ("不合格・補正バッチ", "-", "発生時のみ", ""),
        ],
    },
    {
        "name": "⑦充填・包装",
        "desc": "鉄缶 (錆対策塗装内面) / プラスチック容器 / バッグインボックス。",
        "in": [
            ("合格塗料", "-", "9,640 kg", ""),
            ("空缶・容器・ラベル", "内面ラッカー処理鉄缶 等", "缶数分", "★ 水性は内面防錆必須"),
            ("電力", "-", "数 kWh", ""),
        ],
        "out": [
            ("製品", "出荷形態", "9,600 kg", ""),
            ("滴下塗料・残留", "-", "数十 kg", "産廃"),
            ("包装廃材", "-", "-", "リサイクル"),
        ],
    },
    {
        "name": "並行: 設備洗浄",
        "desc": "水洗が中心。少量の界面活性剤で内部洗浄。廃シンナー発生が少ないのが利点。",
        "in": [
            ("洗浄水", "工業用水", "数 m³/回", ""),
            ("界面活性剤・アルカリ洗剤", "-", "微量", ""),
        ],
        "out": [
            ("洗浄排水 (塗料混入)", "BOD/COD 上昇", "数 m³/回", "排水処理プラント (凝集沈殿)"),
            ("汚泥", "凝集沈殿物", "数 kg", "産廃"),
        ],
    },
],

# ─────────── ③ 粉体塗料 ───────────
"粉体塗料": [
    {
        "name": "①原料配合 (Pre-mixing)",
        "desc": "固体樹脂 (エポキシ/ポリエステル等)、硬化剤、顔料、添加剤、流動性向上剤を粉体ブレンダーで均一混合。",
        "in": [
            ("固体樹脂", "エポキシ/ポリエステル/アクリル", "5,500 kg", "塗膜骨格 (固形)"),
            ("硬化剤", "ジシアンジアミド/TGIC/HAA 等", "500 kg", "架橋反応剤"),
            ("顔料", "TiO₂・無機顔料", "2,500 kg", "着色・隠蔽"),
            ("体質顔料", "炭酸カルシウム/硫酸バリウム", "1,000 kg", "増量・流動性"),
            ("添加剤 (流動性・脱泡)", "ベンゾイン等", "500 kg", ""),
            ("電力 (ブレンダー)", "Scope 2", "微", ""),
        ],
        "out": [
            ("配合粉体", "プリミックス", "10,000 kg", "押出工程へ"),
            ("空袋・粉塵", "PE 袋・TiO₂ 等", "数 kg + 微量", "集塵機・産廃"),
        ],
    },
    {
        "name": "②押出混練 (Extrusion)",
        "desc": "二軸エクストルーダーで 90〜120℃ に加熱しながら混練。樹脂溶融状態で顔料を分散。滞留時間は数十秒〜分のレベルで瞬間的。設備投資 5〜20 億円規模 (業界一般感)。",
        "in": [
            ("配合粉体", "前工程産物", "10,000 kg", ""),
            ("電力 (エクストルーダー 50〜300 kW)", "Scope 2 大消費", "業界一般値の概算", "★最大電力消費工程"),
            ("冷却水 (バレル冷却)", "-", "数 m³/t", "温度暴走防止"),
        ],
        "out": [
            ("溶融シート (チップ)", "押出物", "9,990 kg", "冷却工程へ"),
            ("発熱", "顕熱", "-", "回収可能"),
            ("微量 VOC", "添加剤分解物", "微量", "排気処理"),
        ],
    },
    {
        "name": "③冷却 (Cooling)",
        "desc": "冷却ベルト / 冷却ロールで溶融シートを瞬時に冷却・固化。フレーク状に粗砕。",
        "in": [
            ("溶融シート", "前工程産物", "9,990 kg", ""),
            ("冷却ベルト", "金属ベルト", "-", ""),
            ("冷風", "-", "-", ""),
            ("電力", "Scope 2", "微", ""),
        ],
        "out": [
            ("固化フレーク", "粗砕シート", "9,990 kg", "粉砕工程へ"),
            ("温風", "排熱", "-", "回収検討対象"),
        ],
    },
    {
        "name": "④粉砕 (Pulverizing)",
        "desc": "ジェットミル/ハンマーミル/分級ミルで微粉砕。低融点品は低温窒素冷却下で粉砕。",
        "in": [
            ("固化フレーク", "前工程産物", "9,990 kg", ""),
            ("電力 (粉砕機 30〜100 kW)", "Scope 2", "業界一般値の概算", ""),
            ("低温窒素 (低融点品のみ)", "N₂", "-", "粉砕熱抑制"),
        ],
        "out": [
            ("粗粉砕粉", "粒度バラつき", "9,980 kg", "分級工程へ"),
            ("摩耗粉 (機械由来)", "微量", "微", ""),
            ("騒音・振動", "-", "-", ""),
        ],
    },
    {
        "name": "⑤分級 (Classification / Sieving)",
        "desc": "篩・気流分級機で粒度を揃える。製品粒径 平均 20〜40μm。規格外は配合工程へ戻して再利用。",
        "in": [
            ("粗粉砕粉", "前工程産物", "9,980 kg", ""),
            ("電力 (分級機)", "Scope 2", "微", ""),
        ],
        "out": [
            ("規格内粉 (製品)", "平均粒径 20〜40 μm", "9,500 kg (約 95%)", "製品化"),
            ("粗粒戻し", "規格外大粒径", "数百 kg", "再粉砕 or 配合戻し"),
            ("微粉 (オーバーサイズ)", "粒径過小", "数十 kg", "産廃 or 戻し再利用"),
        ],
    },
    {
        "name": "⑥品質検査",
        "desc": "粒度分布・流動性・ゲル化時間・ガス発生量・塗膜外観試験。",
        "in": [
            ("粉体サンプル", "-", "0.5〜2 kg", ""),
            ("試験設備", "電力", "微", ""),
        ],
        "out": [
            ("合格粉体", "-", "9,500 kg", ""),
            ("QC データ", "-", "-", ""),
        ],
    },
    {
        "name": "⑦充填・包装",
        "desc": "袋 (20kg / 25kg) / フレコン (500〜1000kg) に窒素置換しながら充填。湿気厳禁。",
        "in": [
            ("合格粉体", "-", "9,500 kg", ""),
            ("PE 袋 / フレコン", "防湿仕様", "袋数分", ""),
            ("窒素ガス (置換用)", "N₂", "微量", "湿気・吸湿防止"),
            ("電力", "-", "微", ""),
        ],
        "out": [
            ("製品", "袋詰・フレコン", "9,480 kg", "出荷"),
            ("滴下・残粉", "充填損失", "数 kg", "回収 or 産廃"),
            ("不良袋", "-", "少量", ""),
        ],
    },
    {
        "name": "並行: 設備洗浄",
        "desc": "粉体ブラスト・乾燥洗浄が中心。エクストルーダーは樹脂パージで内部洗浄。VOC・排水ゼロ。",
        "in": [
            ("パージ樹脂", "前ロットと同系統", "数十 kg", "色替え用"),
            ("圧縮空気・電力", "-", "微", "粉吹き飛ばし"),
        ],
        "out": [
            ("混色パージ粉", "規格外混色品", "数十 kg", "産廃 or 黒系・グレー系として再利用"),
            ("廃液・廃シンナーなし", "★液体塗料との最大の違い", "0", "VOC ゼロ"),
        ],
    },
],

# ─────────── ④ 電着塗料 (E-coat) ───────────
"電着塗料": [
    {
        "name": "①樹脂中和・水分散",
        "desc": "アミン変性エポキシ (カチオン) または酸変性ポリエステル (アニオン) を有機酸 (カチオン)/アミン (アニオン) で中和し水中に分散。",
        "in": [
            ("樹脂 (アミン変性エポキシ等)", "中和前固体・液体", "5,000 kg", "塗膜骨格"),
            ("中和剤 (有機酸/アミン)", "酢酸/乳酸 など", "200 kg", "イオン化"),
            ("水 (純水)", "イオン交換水", "4,000 kg", "媒体"),
            ("ブロック化イソシアネート", "硬化剤", "500 kg", "焼付硬化用"),
            ("電力 (撹拌)", "Scope 2", "微", ""),
        ],
        "out": [
            ("樹脂分散液", "Emulsion", "9,700 kg", "次工程へ"),
            ("微量揮散物", "中和剤の一部", "微量", "排気処理"),
        ],
    },
    {
        "name": "②顔料ペースト調製",
        "desc": "顔料分散用樹脂・顔料・水を別ラインで配合し、ビーズミルで分散。",
        "in": [
            ("顔料分散用樹脂", "専用グラインドビヒクル", "300 kg", ""),
            ("顔料 (カーボン・酸化鉄等)", "防錆顔料含む", "700 kg", "★防錆機能付与"),
            ("水", "純水", "200 kg", "媒体"),
            ("電力 (ビーズミル)", "Scope 2", "業界一般値の概算", ""),
        ],
        "out": [
            ("顔料ペースト", "高濃度分散体", "1,200 kg", "次工程へ"),
            ("摩耗ビーズ", "-", "微", "産廃"),
        ],
    },
    {
        "name": "③配合・調整",
        "desc": "樹脂分散液と顔料ペーストを規定比で混合し、最終的に固形分 15〜25% の電着塗料に調整。",
        "in": [
            ("樹脂分散液", "前工程産物", "9,700 kg", ""),
            ("顔料ペースト", "前工程産物", "1,200 kg", ""),
            ("追加水・添加剤", "-", "→ 量に含む", "希釈・性能調整"),
            ("低速撹拌", "Scope 2", "微", ""),
        ],
        "out": [
            ("電着塗料 (調整完了)", "固形分 15〜25%", "10,900 kg", "次工程へ"),
            ("微量揮散物", "-", "微", ""),
        ],
    },
    {
        "name": "④ろ過",
        "desc": "細目フィルタで異物除去。塗膜外観に直結するため厳密管理。",
        "in": [
            ("電着塗料", "前工程産物", "10,900 kg", ""),
            ("フィルター", "PP/メッシュ", "数枚", ""),
        ],
        "out": [
            ("清浄電着塗料", "-", "10,880 kg", ""),
            ("使用済フィルター", "-", "数 kg", "産廃"),
        ],
    },
    {
        "name": "⑤品質検査",
        "desc": "固形分・pH・電気伝導度・粒径・MEQ (メチルエタノアミン当量) などを測定。",
        "in": [
            ("塗料サンプル", "-", "1〜2 kg", ""),
            ("試験設備", "電力", "微", ""),
        ],
        "out": [
            ("合格電着塗料", "-", "10,880 kg", ""),
            ("QC データ", "-", "-", ""),
        ],
    },
    {
        "name": "⑥充填・出荷",
        "desc": "通常はタンクローリー / IBC で出荷。顧客 (自動車 OEM) の塗装ラインタンクに直送が一般的。",
        "in": [
            ("合格塗料", "-", "10,880 kg", ""),
            ("タンクローリー / IBC", "容器", "1 台 〜数台分", ""),
            ("電力 (移送ポンプ)", "Scope 2", "微", ""),
        ],
        "out": [
            ("出荷ロット", "-", "10,850 kg", "OEM 塗装ラインへ"),
            ("配管残液", "微量", "数 kg", "次バッチへ転用 or 廃棄"),
        ],
    },
    {
        "name": "並行: 設備洗浄",
        "desc": "水洗が中心。樹脂残渣は限外ろ過 (UF) で塗料側に戻すリカバリー運用が標準。",
        "in": [
            ("洗浄水 (純水)", "-", "数 m³/回", ""),
            ("電力 (UF 装置)", "Scope 2", "数 kWh", ""),
        ],
        "out": [
            ("洗浄排水", "塗料微量混入", "数 m³/回", "排水処理プラント"),
            ("回収塗料 (UF 透過後)", "-", "数 kg", "再投入"),
        ],
    },
],

# ─────────── ⑤ UV硬化塗料 ───────────
"UV硬化塗料": [
    {
        "name": "①計量・配合 (遮光下)",
        "desc": "オリゴマー・モノマー・光開始剤・添加剤を計量。光開始剤の感光を防ぐため遮光環境 (黄色灯・遮光容器) で作業。",
        "in": [
            ("ウレタンアクリレート等オリゴマー", "主バインダー", "5,000 kg", "塗膜形成主成分"),
            ("反応性モノマー", "TPGDA/HDDA 等", "3,000 kg", "粘度低減・反応希釈剤"),
            ("光開始剤", "Irgacure 等", "300 kg", "★UV で活性ラジカル発生"),
            ("顔料・添加剤", "クリアの場合は微量", "1,000 kg", "色調・流動性"),
            ("シリコン系添加剤", "BYK 等", "微量", "表面性向上"),
            ("電力", "Scope 2", "微", "計量設備"),
        ],
        "out": [
            ("配合バッチ", "感光性", "9,300 kg", "次工程へ (遮光)"),
            ("空袋・空ドラム", "-", "数十 kg", ""),
        ],
    },
    {
        "name": "②予備混合・分散",
        "desc": "ディスパー・ビーズミルで分散。UV 塗料は溶剤レスのため発熱しやすく、冷却制御が重要。",
        "in": [
            ("配合バッチ", "前工程産物", "9,300 kg", ""),
            ("電力 (分散機)", "Scope 2", "業界一般値の概算", ""),
            ("冷却水", "強冷却必要", "数 m³/t", "★発熱大"),
        ],
        "out": [
            ("ミルベース (感光性)", "粒径規格内", "9,280 kg", "遮光下で次工程へ"),
            ("発熱", "顕熱", "-", "冷却水で除去"),
        ],
    },
    {
        "name": "③レットダウン",
        "desc": "追加モノマー・光開始剤を遮光下で添加し最終調整。",
        "in": [
            ("ミルベース", "前工程産物", "9,280 kg", ""),
            ("追加モノマー・添加剤", "-", "→ 量に含む", ""),
            ("低速撹拌", "Scope 2", "微", ""),
        ],
        "out": [
            ("UV 硬化塗料 (ベース)", "感光性", "9,300 kg", "調色 or ろ過へ"),
        ],
    },
    {
        "name": "④調色 (有色品のみ)",
        "desc": "UV 用着色剤で CCM 調色。光開始剤の吸収波長と顔料の遮光性を考慮 (光が届かないと硬化しない)。",
        "in": [
            ("ベース塗料", "-", "9,300 kg", ""),
            ("UV 用着色剤", "-", "10〜200 kg", ""),
        ],
        "out": [
            ("有色 UV 塗料", "-", "9,400 kg", ""),
            ("試験塗板", "-", "数百 g", "産廃"),
        ],
    },
    {
        "name": "⑤ろ過 (遮光下)",
        "desc": "細目フィルタで凝集物除去。光遮蔽配管・フィルタハウジング使用。",
        "in": [
            ("塗料", "-", "9,400 kg", ""),
            ("フィルター (遮光仕様)", "-", "数枚", ""),
        ],
        "out": [
            ("清浄 UV 塗料", "-", "9,390 kg", ""),
            ("使用済フィルター", "-", "数 kg", "産廃"),
        ],
    },
    {
        "name": "⑥品質検査",
        "desc": "粘度・色差に加え、UV 硬化試験 (照射量と硬化深度)・ゲル化時間を測定。",
        "in": [
            ("塗料サンプル", "-", "1〜2 kg", ""),
            ("UV 試験機", "電力", "微", ""),
        ],
        "out": [
            ("合格塗料", "-", "9,390 kg", ""),
            ("QC データ", "-", "-", ""),
        ],
    },
    {
        "name": "⑦充填 (遮光容器)",
        "desc": "黒色容器・遮光バッグインボックスに充填。透明缶は使用不可。",
        "in": [
            ("合格塗料", "-", "9,390 kg", ""),
            ("遮光容器・ラベル", "黒色プラ/金属缶", "缶数分", ""),
            ("電力", "-", "微", ""),
        ],
        "out": [
            ("製品 (遮光梱包)", "出荷形態", "9,360 kg", "出荷"),
            ("滴下塗料", "充填損失", "数十 kg", "産廃 (★感光性のため早期処理)"),
            ("不良缶", "-", "少量", ""),
        ],
    },
    {
        "name": "並行: 設備洗浄",
        "desc": "モノマーが洗浄液を兼ねる場合あり。完全遮光下で実施。",
        "in": [
            ("洗浄モノマー or 専用洗浄剤", "-", "50〜200 kg/回", ""),
            ("電力", "-", "微", ""),
        ],
        "out": [
            ("廃モノマー (特管産廃)", "感光性", "50〜200 kg/回", "焼却 (再生困難)"),
            ("VOC ほぼゼロ", "★溶剤レスの強み", "0", "環境性能高"),
        ],
    },
],

}


# ────────────────────────────────────────────────────────────────────
# シート生成関数
# ────────────────────────────────────────────────────────────────────

def setup_columns(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_cell(ws, row, col, value, font=None, fill=None, align=None, border=BORDER_ALL):
    c = ws.cell(row=row, column=col, value=value)
    if font: c.font = font
    if fill: c.fill = fill
    if align: c.alignment = align
    if border: c.border = border
    return c


def make_summary_sheet(wb):
    ws = wb.create_sheet("表紙_目次", 0)
    setup_columns(ws, [4, 28, 35, 35, 35])

    ws.merge_cells("B2:E2")
    write_cell(ws, 2, 2, "塗料メーカー 製造フロー & 工程別 IN/OUT 表",
               font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 30

    ws.merge_cells("B3:E3")
    write_cell(ws, 3, 2, "テンプレート出典: 添付 PDF「リン酸(湿式法) IN/OUT表(マテリアルバランス)」を踏襲",
               font=Font(name="Yu Gothic", size=9, italic=True),
               align=ALIGN_CENTER, border=None)

    # 目次
    write_cell(ws, 5, 2, "シート構成", font=FONT_H, fill=FILL_HEADER_PROC, align=ALIGN_CENTER)
    ws.merge_cells("B5:E5")

    rows = [
        ("①", "製品タイプ一覧",        "塗料の主要 5 種類について用途・特徴・規制を整理"),
        ("②", "製造フロー比較",        "5 種類の製造フローを工程順に並べて比較"),
        ("③-1", "溶剤系液体塗料 IN/OUT", "汎用・自動車補修・工業用の標準ライン"),
        ("③-2", "水性液体塗料 IN/OUT",   "建築用主体。VOC 削減型"),
        ("③-3", "粉体塗料 IN/OUT",       "家電・自動車部品。VOC ゼロ"),
        ("③-4", "電着塗料 IN/OUT",       "自動車車体下塗。タンクローリー出荷"),
        ("③-5", "UV硬化塗料 IN/OUT",     "木工・印刷・プラ。遮光下で製造"),
        ("出典", "参考文献",             "JPMA・環境省・経産省・装置メーカー資料 等"),
    ]
    write_cell(ws, 6, 2, "番号", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    write_cell(ws, 6, 3, "シート名", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    ws.merge_cells("D6:E6")
    write_cell(ws, 6, 4, "内容", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    for i, (no, name, desc) in enumerate(rows, start=7):
        write_cell(ws, i, 2, no, font=FONT_BODY, align=ALIGN_CENTER)
        write_cell(ws, i, 3, name, font=FONT_BODY, align=ALIGN_LEFT)
        ws.merge_cells(f"D{i}:E{i}")
        write_cell(ws, i, 4, desc, font=FONT_BODY, align=ALIGN_LEFT)

    # 凡例
    base = 7 + len(rows) + 2
    ws.merge_cells(f"B{base}:E{base}")
    write_cell(ws, base, 2, "テンプレート凡例 (リン酸 PDF と同形式)",
               font=FONT_H, fill=FILL_HEADER_PROC, align=ALIGN_CENTER)
    legend = [
        ("上段 (IN)", "工程に投入される原料・ユーティリティ", FILL_IN),
        ("中段 (工程)", "工程名・装置・概要", FILL_PROCESS),
        ("下段 (OUT)", "工程から出る中間品・廃棄物・排出", FILL_OUT),
    ]
    for i, (k, v, fill) in enumerate(legend, start=base + 1):
        write_cell(ws, i, 2, k, font=FONT_H, fill=fill, align=ALIGN_CENTER)
        ws.merge_cells(f"C{i}:E{i}")
        write_cell(ws, i, 3, v, font=FONT_BODY, align=ALIGN_LEFT)


def make_product_type_sheet(wb):
    ws = wb.create_sheet("①製品タイプ一覧")
    setup_columns(ws, [3, 6, 22, 30, 38, 38])

    ws.merge_cells("B2:F2")
    write_cell(ws, 2, 2, "① 塗料メーカーにおける主要製品タイプ",
               font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 28

    headers = ["No.", "製品タイプ", "用途代表例", "特徴", "関連規制・留意点"]
    for i, h in enumerate(headers, start=2):
        write_cell(ws, 4, i, h, font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    ws.row_dimensions[4].height = 22

    for i, pt in enumerate(PRODUCT_TYPES, start=5):
        write_cell(ws, i, 2, pt["id"], font=FONT_BODY, align=ALIGN_CENTER)
        write_cell(ws, i, 3, pt["name"], font=Font(name="Yu Gothic", size=10, bold=True), align=ALIGN_LEFT)
        write_cell(ws, i, 4, pt["用途代表例"], font=FONT_BODY, align=ALIGN_WRAP)
        write_cell(ws, i, 5, pt["特徴"], font=FONT_BODY, align=ALIGN_WRAP)
        write_cell(ws, i, 6, pt["規制"], font=FONT_BODY, align=ALIGN_WRAP)
        ws.row_dimensions[i].height = 45


def make_flow_compare_sheet(wb):
    ws = wb.create_sheet("②製造フロー比較")
    # 製品種5列 × 工程
    setup_columns(ws, [3, 5, 30, 30, 30, 30, 30])

    ws.merge_cells("B2:G2")
    write_cell(ws, 2, 2, "② 製品タイプ別 製造フロー比較",
               font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 28

    write_cell(ws, 4, 2, "工程#", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    for i, pt in enumerate(PRODUCT_TYPES, start=3):
        write_cell(ws, 4, i, pt["name"], font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    ws.row_dimensions[4].height = 24

    # 各製品の工程名を縦に並べる
    max_steps = max(len(FLOWS[pt["name"]]) for pt in PRODUCT_TYPES)
    for step_i in range(max_steps):
        row = 5 + step_i
        write_cell(ws, row, 2, f"#{step_i+1}", font=FONT_H, align=ALIGN_CENTER, fill=FILL_PROCESS)
        for col_i, pt in enumerate(PRODUCT_TYPES, start=3):
            flow = FLOWS[pt["name"]]
            text = flow[step_i]["name"] if step_i < len(flow) else "-"
            write_cell(ws, row, col_i, text, font=FONT_BODY, align=ALIGN_WRAP)
        ws.row_dimensions[row].height = 35


def make_inout_sheet(wb, product_name, idx):
    """1製品タイプ分の IN/OUT 表シートを作成。

    レイアウト (リン酸 PDF と同形式):
      列  B   = ラベル (IN / 工程 / OUT)
      列  C,D,E,F = 工程1ブロック (品名/標準成分/量/備考)
      列  G,H,I,J = 工程2ブロック ...
      ...
      行構造:
      タイトル
      工程名ヘッダ行
      IN ヘッダ (品名/標準成分/量/備考) × 工程数
      IN データ行 × N (最大IN品目数に合わせる)
      工程説明行
      OUT ヘッダ (同上)
      OUT データ行 × M
    """
    sheet_name = f"③-{idx} {product_name}"
    if len(sheet_name) > 31:
        sheet_name = sheet_name[:31]
    ws = wb.create_sheet(sheet_name)

    flow = FLOWS[product_name]
    n_steps = len(flow)

    # 列幅: B=8, 各工程の4列 = 22, 18, 12, 28
    widths = [3, 8]
    for _ in range(n_steps):
        widths.extend([22, 18, 12, 28])
    setup_columns(ws, widths)

    last_col_idx = 2 + 1 + n_steps * 4  # 1始まり最後の列
    last_col = get_column_letter(last_col_idx)

    # タイトル
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col_idx)
    write_cell(ws, 2, 2, f"③-{idx}  {product_name}  IN/OUT 表 (マテリアルバランス)",
               font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 28

    # 凡例行 (注記)
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=last_col_idx)
    write_cell(ws, 3, 2,
               "テンプレート出典: 添付 PDF「リン酸(湿式法) IN/OUT表」を踏襲。  "
               "量は 10t バッチを基準とした業界一般値の概算 (工場・製品・処方で大きく変動)。",
               font=Font(name="Yu Gothic", size=8, italic=True, color="555555"),
               align=ALIGN_CENTER, border=None)
    ws.row_dimensions[3].height = 18

    # 工程名ヘッダ (5行目)
    row_step_header = 5
    write_cell(ws, row_step_header, 2, "工程", font=FONT_H, fill=FILL_HEADER_PROC, align=ALIGN_CENTER)
    for s_i, step in enumerate(flow):
        c0 = 3 + s_i * 4
        ws.merge_cells(start_row=row_step_header, start_column=c0,
                       end_row=row_step_header, end_column=c0 + 3)
        write_cell(ws, row_step_header, c0, step["name"],
                   font=FONT_PROC, fill=FILL_HEADER_PROC, align=ALIGN_CENTER)
    ws.row_dimensions[row_step_header].height = 30

    # IN ヘッダ (6行目)
    row_in_header = 6
    write_cell(ws, row_in_header, 2, "IN", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    sub_headers = ["品名", "標準成分", "量", "備考/役割"]
    for s_i in range(n_steps):
        for k, sh in enumerate(sub_headers):
            write_cell(ws, row_in_header, 3 + s_i * 4 + k, sh,
                       font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    ws.row_dimensions[row_in_header].height = 20

    # IN データ行: 最大IN品目数を取る
    max_in = max(len(s["in"]) for s in flow)
    for i in range(max_in):
        r = row_in_header + 1 + i
        write_cell(ws, r, 2, f"IN-{i+1}", font=FONT_BODY,
                   fill=FILL_IN, align=ALIGN_CENTER)
        for s_i, step in enumerate(flow):
            items = step["in"]
            if i < len(items):
                vals = items[i]
            else:
                vals = ("", "", "", "")
            for k, v in enumerate(vals):
                write_cell(ws, r, 3 + s_i * 4 + k, v,
                           font=FONT_BODY, fill=FILL_IN, align=ALIGN_WRAP)
        ws.row_dimensions[r].height = 30

    # 工程説明行
    row_desc = row_in_header + 1 + max_in
    write_cell(ws, row_desc, 2, "工程説明", font=FONT_H, fill=FILL_HEADER_PROC, align=ALIGN_CENTER)
    for s_i, step in enumerate(flow):
        c0 = 3 + s_i * 4
        ws.merge_cells(start_row=row_desc, start_column=c0,
                       end_row=row_desc, end_column=c0 + 3)
        write_cell(ws, row_desc, c0, step["desc"],
                   font=FONT_BODY, fill=FILL_PROCESS, align=ALIGN_WRAP)
    ws.row_dimensions[row_desc].height = 90

    # OUT ヘッダ
    row_out_header = row_desc + 1
    write_cell(ws, row_out_header, 2, "OUT", font=FONT_H, fill=FILL_HEADER_OUT, align=ALIGN_CENTER)
    for s_i in range(n_steps):
        for k, sh in enumerate(sub_headers):
            write_cell(ws, row_out_header, 3 + s_i * 4 + k, sh,
                       font=FONT_H, fill=FILL_HEADER_OUT, align=ALIGN_CENTER)
    ws.row_dimensions[row_out_header].height = 20

    # OUT データ行
    max_out = max(len(s["out"]) for s in flow)
    for i in range(max_out):
        r = row_out_header + 1 + i
        write_cell(ws, r, 2, f"OUT-{i+1}", font=FONT_BODY,
                   fill=FILL_OUT, align=ALIGN_CENTER)
        for s_i, step in enumerate(flow):
            items = step["out"]
            if i < len(items):
                vals = items[i]
            else:
                vals = ("", "", "", "")
            for k, v in enumerate(vals):
                write_cell(ws, r, 3 + s_i * 4 + k, v,
                           font=FONT_BODY, fill=FILL_OUT, align=ALIGN_WRAP)
        ws.row_dimensions[r].height = 30

    # 印刷設定: 横向き
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = "C6"


def make_sources_sheet(wb):
    ws = wb.create_sheet("出典_参考文献")
    setup_columns(ws, [3, 6, 50, 55])

    ws.merge_cells("B2:D2")
    write_cell(ws, 2, 2, "出典・参考文献 (Web 整合確認に使用)",
               font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 28

    write_cell(ws, 4, 2, "No.", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    write_cell(ws, 4, 3, "情報源", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)
    write_cell(ws, 4, 4, "URL", font=FONT_H, fill=FILL_HEADER_IN, align=ALIGN_CENTER)

    sources = [
        ("日本塗料工業会 (JPMA) VOC規制関連情報", "https://www.toryo.or.jp/jp/anzen/VOC/index.html"),
        ("JPMA 塗料からのVOC排出実態推計のまとめ 2022年度", "https://www.toryo.or.jp/jp/book/voch2022.html"),
        ("JPMA コーティング・ケア環境管理指標", "https://www.toryo.or.jp/jp/anzen/cc/03_feature.html"),
        ("関東塗料工業組合 2022年度塗料VOC排出推計", "https://kantoko.com/blog/2024/04/90578/"),
        ("庫本「塗料における分散」色材協会誌 (J-Stage)",
         "https://www.jstage.jst.go.jp/article/shikizai1937/78/4/78_191/_pdf/-char/ja"),
        ("塗料における最近の顔料分散とそのプロセス (J-Stage)",
         "https://www.jstage.jst.go.jp/article/shikizai/87/6/87_204/_pdf"),
        ("Mixing Tank 塗料製造の方法",
         "https://www.mixing-tank.com/ja/blog/"),
        ("JCT Machinery Paint Manufacturing Process",
         "https://www.mixmachinery.com/news/paint-manufacturing-process-jct-machinery.html"),
        ("アシザワ・ファインテック ビーズミルガイド", "https://ashizawa.com/guidance/05.html"),
        ("関西ペイント R&D 高度顔料分散システム",
         "https://asset.kansai.co.jp/uploads/rd/paint_study/pdf/143/04.pdf"),
        ("サンノプコ 顔料分散解説",
         "https://www.sannopco.co.jp/feature/pigment-dispersion/"),
        ("ヒバラコーポレーション 粉体塗装", "https://kougyoutosou.com/technology/powder/"),
        ("日鉄防食 粉体塗装", "https://acc.nipponsteel.com/powder-coating/"),
        ("NCC 粉なのに塗料？", "https://ncc-3clab.com/resolution/painting/2461/"),
        ("NCC 塗着効率とは", "https://ncc-nice.com/ncc-coating/knowledge/basics/coatingefficiency/"),
        ("モノタロウ 粉体塗料解説",
         "https://www.monotaro.com/note/readingseries/tosouqa/0416/"),
        ("e-Gov 危険物の規制に関する政令", "https://laws.e-gov.go.jp/law/334CO0000000306"),
        ("三協化学 消防法と有機溶剤・指定数量", "https://www.sankyo-chem.com/news/post-689/"),
        ("化研テック 指定数量とは",
         "https://www.kaken-tech.co.jp/trouble/%E6%8C%87%E5%AE%9A%E6%95%B0%E9%87%8F/"),
        ("e-reverse.com 廃塗料の廃棄解説", "https://www.e-reverse.com/blog/law067/"),
        ("丸商 シンナーの廃棄物処理",
         "https://marusho-eco.jp/column/how_to_dispose_of_thinner_waste/"),
        ("ネクストリー 溶剤回収装置事例", "https://nextry.jp/340/"),
        ("コスモエンジニアリング パイプライニング",
         "https://www.cosmoeng.co.jp/service/ctg05/piping/pipelining.html"),
        ("環境省 廃棄物処理に関する統計",
         "https://www.env.go.jp/recycle/waste/wastetoukei_index.html"),
        ("経産省 PRTR 塗料に係る排出量",
         "https://www.meti.go.jp/policy/chemical_management/law/prtr/r4kohyo/05todokedegaiyou/syousai/5.pdf"),
    ]
    for i, (title, url) in enumerate(sources, start=5):
        write_cell(ws, i, 2, i - 4, font=FONT_BODY, align=ALIGN_CENTER)
        write_cell(ws, i, 3, title, font=FONT_BODY, align=ALIGN_LEFT)
        c = write_cell(ws, i, 4, url, font=FONT_BODY, align=ALIGN_LEFT)
        c.hyperlink = url
        c.font = Font(name="Yu Gothic", size=9, color="0563C1", underline="single")
        ws.row_dimensions[i].height = 22

    # 注記
    note_row = 5 + len(sources) + 2
    ws.merge_cells(start_row=note_row, start_column=2, end_row=note_row, end_column=4)
    write_cell(ws, note_row, 2,
               "注: 工程別の電力 (kWh/t)・歩留まり・廃棄物単価などの定量値は業界一般値の概算であり、"
               "工場・製品種・年度で大きく変動。実プロジェクト適用時は対象工場の実測値および JPMA 公表値で上書き校正のこと。",
               font=Font(name="Yu Gothic", size=8, italic=True, color="555555"),
               align=ALIGN_WRAP, border=None)
    ws.row_dimensions[note_row].height = 40


# ────────────────────────────────────────────────────────────────────
# メイン
# ────────────────────────────────────────────────────────────────────
def main():
    wb = Workbook()
    # default sheet 削除
    wb.remove(wb.active)

    make_summary_sheet(wb)
    make_product_type_sheet(wb)
    make_flow_compare_sheet(wb)
    for i, pt in enumerate(PRODUCT_TYPES, start=1):
        make_inout_sheet(wb, pt["name"], i)
    make_sources_sheet(wb)

    wb.save(OUTPUT_PATH)
    print(f"OK: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
