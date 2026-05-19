#!/usr/bin/env python3
"""塗料メーカー ヒアリング先 候補企業リスト Excel 生成スクリプト (50社版)

方針 (確定済):
- 対象     : 塗料メーカーのみ (業界団体・装置メーカーは除外)
- 社数     : 50 社
- アプローチ: 網羅型
- 立場     : 業界関係者 (取引先・関連事業)
- 成果物   : リスト Excel のみ
- 検証深さ : 定性 + 定量

シート構成:
  1. 表紙・運用方針
  2. ヒアリング先 候補リスト (メイン、50社)
  3. 優先度スコアリング基準
  4. カテゴリ別カバレッジ確認

データの注意:
  ・売上・従業員数は公開情報ベースの推定値を含む (要校正)
  ・「公開情報充実度」列の "(要確認)" は基本情報の出典再確認が必要なエントリ
  ・実ヒアリング前に貴社情報で必ず校正する前提
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT = "塗料メーカー_ヒアリング候補リスト.xlsx"

thin = Side(style="thin", color="666666")
medium = Side(style="medium", color="333333")
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
BORDER_BOX = Border(left=medium, right=medium, top=medium, bottom=medium)

FILL_TITLE = PatternFill("solid", fgColor="1F3864")
FILL_HDR1 = PatternFill("solid", fgColor="9DC3E6")
FILL_HDR2 = PatternFill("solid", fgColor="A9D08E")
FILL_HDR3 = PatternFill("solid", fgColor="F4B084")

CAT_FILL = {
    "大手総合":         PatternFill("solid", fgColor="FFE699"),
    "中堅総合":         PatternFill("solid", fgColor="FFF2CC"),
    "建築・仕上塗材":   PatternFill("solid", fgColor="E2EFDA"),
    "自動車補修":       PatternFill("solid", fgColor="DDEBF7"),
    "防食・船舶":       PatternFill("solid", fgColor="BDD7EE"),
    "機能・特殊":       PatternFill("solid", fgColor="FCE4D6"),
    "UV・印刷インキ":   PatternFill("solid", fgColor="E4DFEC"),
    "DIY・小口":        PatternFill("solid", fgColor="EDEDED"),
    "日ペグループ":     PatternFill("solid", fgColor="FFD966"),
    "関ペグループ":     PatternFill("solid", fgColor="F8CBAD"),
    "外資系・日本法人": PatternFill("solid", fgColor="C9C9FF"),
}

FONT_TITLE = Font(name="Yu Gothic", size=14, bold=True, color="FFFFFF")
FONT_H = Font(name="Yu Gothic", size=10, bold=True)
FONT_BODY = Font(name="Yu Gothic", size=9)
FONT_NUM = Font(name="Yu Gothic", size=10, bold=True, color="C00000")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def setup_columns(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def cell(ws, r, c, v, font=None, fill=None, align=None, border=BORDER_ALL):
    cc = ws.cell(row=r, column=c, value=v)
    if font: cc.font = font
    if fill: cc.fill = fill
    if align: cc.alignment = align
    if border: cc.border = border
    return cc


HEADERS = [
    "No.", "カテゴリ", "企業名", "過去契約有無", "本社所在地", "上場区分",
    "売上規模(連結)", "従業員数(概数)", "主要セグメント",
    "溶剤系", "水性", "粉体", "電着", "UV硬化",
    "主要工場(代表)", "Web", "IR/問合せ窓口",
    "公開情報充実度", "ヒアリング推奨工程", "確認したい定量項目",
    "優先度スコア", "推奨アプローチルート", "既存関係(想定)",
    "留意点・ハードル", "進捗ステータス", "次アクション", "メモ",
]

# 過去契約有無 (取引明細CSV 2026/05/11 時点との突合結果)
# 親会社のみ厳密マッチで、子会社ヒットは親会社には伝播させない
PAST_CONTRACT = {
    "関西ペイント": 1,
    "ロックペイント": 4,
    "ナトコ": 1,
    "シンロイヒ": 1,
    "大日精化工業": 1,
    "サカタインクス": 4,
    "トウペ": 1,
    "日本ペイント・インダストリアルコーティングス": 3,
}


def C(cat, name, hq, listed, sales, emp, seg,
      solv, water, powder, ed, uv,
      plant, web, ir, info,
      proc, quant, score, route, rel, caveat, memo):
    return {
        "カテゴリ": cat, "企業名": name, "本社所在地": hq, "上場区分": listed,
        "売上規模(連結)": sales, "従業員数(概数)": emp, "主要セグメント": seg,
        "溶剤系": solv, "水性": water, "粉体": powder, "電着": ed, "UV硬化": uv,
        "主要工場(代表)": plant, "Web": web, "IR/問合せ窓口": ir,
        "公開情報充実度": info,
        "ヒアリング推奨工程": proc, "確認したい定量項目": quant,
        "優先度スコア": score, "推奨アプローチルート": route,
        "既存関係(想定)": rel, "留意点・ハードル": caveat,
        "進捗ステータス": "", "次アクション": "", "メモ": memo,
    }


COMPANIES = [
    # ─── A. 国内大手 (2) ───────────────────────────────────────────
    C("大手総合", "日本ペイントホールディングス", "大阪市北区", "東証プライム",
      "約1.5兆円(2024)", "約40,000人(連結)", "汎用・自動車・工業・船舶 (世界4位級)",
      "◎", "◎", "○", "◎", "○",
      "大阪・東京・愛知 ほか国内10数拠点", "https://www.nipponpaint-holdings.com/", "IR部",
      "高 (統合報告書・サステナ報告書)",
      "①〜⑨全工程 + 並行洗浄・排ガス処理",
      "kWh/t・歩留・廃シンナー量・VOC・Scope1/2/3", 95,
      "IR経由 → 技術本部/R&Dセンター", "業界団体・共同セミナー",
      "上場大手のため機密の線引き厳格・NDA必須",
      "M&A 主導、海外拠点も検証視点に"),
    C("大手総合", "関西ペイント", "大阪市中央区", "東証プライム",
      "約5,500億円(2024)", "約15,000人(連結)", "自動車OEM・補修・工業・船舶・電着",
      "◎", "◎", "◎", "◎", "○",
      "大阪・滋賀・茨城(姉ヶ崎) ほか", "https://www.kansai.co.jp/", "IR部/広報部",
      "高 (統合報告書・ESG資料)",
      "①〜⑨ + 電着・粉体の特殊工程",
      "電着MEQ/UF回収率、粉体歩留、CCM導入率", 95,
      "技術部門直接 + IR", "業界団体・装置メーカー紹介",
      "海外子会社(印・南ア)情報の範囲を要相談",
      "EV用電池絶縁の R&D 動向も併せて確認可"),

    # ─── B. 国内中堅・準大手 (8) ──────────────────────────────────
    C("中堅総合", "大日本塗料 (DNT)", "東京都品川区", "東証スタンダード",
      "約750億円(2024)", "約1,800人(連結)", "防食・構造物・蛍光塗料・粉体",
      "◎", "○", "○", "△", "△",
      "大阪・東京・愛知(小牧)", "https://www.dnt.co.jp/", "IR室",
      "中〜高 (有報・統合報告書)",
      "厚膜防食の③④調色・⑦検査・⑧充填",
      "厚塗り防食の固形分・粘度、ジンクリッチプライマー", 82,
      "営業・技術 / IR", "インフラ顧客経由",
      "構造物向け特殊配合は機密性高",
      "ジンクリッチ・蛍光塗料のニッチ例"),
    C("中堅総合", "中国塗料 (CMP)", "東京都千代田区", "東証プライム",
      "約1,100億円(2024)", "約2,500人(連結)", "船舶用塗料(世界4強)・工業防食",
      "◎", "○", "△", "-", "-",
      "兵庫・愛媛 ほか", "https://www.cmp.co.jp/", "IR部",
      "高",
      "船舶塗料の③④(低摩擦・防汚機能の分散)",
      "防汚剤添加量・厚膜塗装の充填精度", 88,
      "IR / 技術センター", "造船所経由",
      "防汚剤組成は競争力源・機密性高",
      "船舶・防食の世界トップ"),
    C("中堅総合", "神東塗料", "兵庫県尼崎市", "東証スタンダード",
      "約280億円", "約700人(連結)", "鉄道車両・工業・電着",
      "◎", "○", "△", "○", "△",
      "尼崎・千葉", "https://www.shintopaint.co.jp/", "総務・IR担当",
      "中",
      "鉄道用塗装ライン①〜⑨、電着塗料の③配合",
      "鉄道塗料の歩留、電着塗料の電気伝導度管理", 76,
      "技術部門 / IR", "鉄道車両メーカー経由",
      "鉄道顧客の機密 (色番号等) は除外",
      "鉄道車両用塗料のリファレンス"),
    C("中堅総合", "日本特殊塗料 (NTC)", "東京都北区", "東証スタンダード",
      "約580億円", "約1,800人(連結)", "建築防音材・自動車防音・特殊塗料",
      "◎", "○", "△", "-", "△",
      "東京・愛知・滋賀 ほか", "https://www.ntpaint.co.jp/", "IR室",
      "中〜高",
      "建築防水塗料・自動車制振塗料の③④⑤",
      "厚膜塗料の塗布量、車両用制振材の充填", 78,
      "IR / 技術部門", "建築・自動車業界経由",
      "防音材事業との切り分けが必要",
      "防音・制振との複合用途が特徴"),
    C("中堅総合", "エスケー化研 (SK化研)", "大阪府茨木市", "非上場(独立系)",
      "約1,800億円(推定)", "約1,400人", "建築用塗料 国内首位",
      "○", "◎", "△", "-", "-",
      "大阪・茨木 ほか全国", "https://www.sk-kaken.co.jp/", "本社・各支店",
      "中 (非上場のためIR限定的)",
      "水性塗料 ②③④調色 + ⑥ろ過 + 設備洗浄",
      "水性ベースの歩留、CCM普及率、防腐剤添加量", 88,
      "技術部門 / 営業窓口", "業界団体 (JPMA) 経由が確実",
      "非上場のため定量データ開示範囲を要相談",
      "建築用水性塗料のリファレンスとして最重要"),
    C("中堅総合", "ロックペイント", "大阪市西淀川区", "非上場",
      "約220億円規模(推定)", "約700人", "自動車補修・建築・工業",
      "◎", "○", "-", "-", "-",
      "大阪・名古屋", "https://www.rockpaint.co.jp/", "営業本部",
      "中",
      "2液型補修塗料の④⑤調色・⑧充填",
      "2液型のポットライフ管理・小ロットCCM導入率", 78,
      "営業 / 補修工場経由", "ディーラー・板金工場経由",
      "色見本データは機密性高",
      "自動車補修の代表的リファレンス"),
    C("中堅総合", "武蔵塗料ホールディングス", "東京都北区", "非上場",
      "約300億円規模(推定)", "約800人(連結)", "工業用・電子部品・粉体・海外展開",
      "○", "○", "◎", "△", "○",
      "栃木・東京・海外(中国・東南アジア)", "https://www.musashipaint.co.jp/", "本社",
      "中",
      "粉体塗料の①〜⑦, 海外工場比較",
      "押出機消費電力・粉砕粒度管理・分級ロス率", 80,
      "技術部門 / 海外拠点", "家電・電子業界経由",
      "海外グループの統一データ取得は難",
      "粉体塗料の代表ヒアリング先"),
    C("中堅総合", "立川塗料", "兵庫県川西市", "非上場",
      "約100億円規模(推定)", "約300人", "工業用・自動車部品・建設機械",
      "◎", "○", "△", "-", "-",
      "川西・三田", "https://www.tachikawatoryo.co.jp/", "本社窓口",
      "低 (非上場・コーポレートサイトのみ)",
      "工業用塗料の①計量〜⑧充填",
      "中ロット工業塗料の歩留・在庫回転", 65,
      "業界団体経由・取引先紹介", "建機メーカー経由",
      "中小規模のため工場見学受入余力要確認",
      "中堅独立系の代表例"),

    # ─── C. 自動車補修 (2) ─────────────────────────────────────
    C("自動車補修", "イサム塗料", "大阪市住之江区", "非上場",
      "約160億円規模(推定)", "約450人", "自動車補修・工業",
      "◎", "△", "-", "-", "-",
      "大阪・茨城", "https://www.isamu.co.jp/", "本社問合せ",
      "低〜中",
      "補修塗料の④⑤調色・⑥ろ過",
      "小ロット調色のサンプル廃棄率", 70,
      "営業・既存取引", "板金塗装業界経由",
      "ロックペイントと役割が重なる場合あり",
      "ロックペイントとの比較対象"),
    C("自動車補修", "ホーマー技研", "東京都品川区", "非上場",
      "数十億円規模(推定)", "約100人(推定)", "自動車補修副資材・特殊塗料",
      "○", "△", "-", "-", "-",
      "(要確認)", "https://www.homer-jp.com/", "本社問合せ",
      "低 (要確認)",
      "補修副資材の充填工程",
      "副資材の SKU 数・小ロット運用", 55,
      "営業窓口", "板金塗装業界経由",
      "塗料本体ではなく補修副資材寄りの可能性",
      "周辺領域として参考"),

    # ─── D. 建築・仕上塗材 (5) ─────────────────────────────────
    C("建築・仕上塗材", "スズカファイン", "三重県鈴鹿市", "非上場",
      "約230億円規模(推定)", "約600人", "建築仕上塗材・水性塗料",
      "○", "◎", "△", "-", "-",
      "鈴鹿・関東", "https://www.suzukafine.co.jp/", "本社問合せ",
      "中",
      "仕上塗材の②分散・⑤調色",
      "仕上塗材の粒径管理・水性化率", 76,
      "技術部門", "ゼネコン経由",
      "仕上塗材は塗料分類で特殊な扱い",
      "建築仕上塗材の代表"),
    C("建築・仕上塗材", "菊水化学工業", "名古屋市中区", "東証スタンダード",
      "約220億円", "約500人(連結)", "建築仕上塗材・防水・耐火塗料",
      "○", "◎", "△", "-", "-",
      "愛知・茨城", "https://www.kikusui-chem.co.jp/", "IR担当",
      "中〜高",
      "仕上塗材・防水材の③④⑤",
      "防水材の固形分管理、耐火塗料の塗膜厚", 74,
      "IR / 技術部門", "ゼネコン経由",
      "耐火塗料は機密性高",
      "上場の建築特化メーカー"),
    C("建築・仕上塗材", "フッコー", "兵庫県神戸市", "非上場",
      "約60億円規模(推定)", "約200人", "建築仕上塗材",
      "△", "◎", "△", "-", "-",
      "兵庫・関東", "https://www.fukko.co.jp/", "本社問合せ",
      "低〜中",
      "仕上塗材の②分散・⑤調色",
      "仕上塗材中小プラントの規模感", 58,
      "営業窓口", "建材販売店経由",
      "中小規模ゆえスケール感に注意",
      "中小建築塗材の補強"),
    C("建築・仕上塗材", "アトミクス", "東京都荒川区", "非上場",
      "約180億円規模(推定)", "約500人", "床用塗料・建築・水性",
      "◎", "◎", "-", "-", "△",
      "茨城・千葉", "https://www.atomix.co.jp/", "本社問合せ",
      "低〜中",
      "水性床塗料の②分散・⑤調色",
      "床用塗料の固形分・耐摩耗性管理", 68,
      "営業・技術窓口", "ゼネコン経由",
      "床塗料特化のため一般用との差異に注意",
      "床塗料のニッチ事例"),
    C("建築・仕上塗材", "大同塗料", "大阪市西成区", "非上場",
      "約100億円規模(推定)", "約300人", "建築・橋梁・構造物",
      "◎", "○", "△", "-", "-",
      "大阪・茨城", "https://www.daidotoryo.co.jp/", "本社問合せ",
      "中",
      "重防食塗料の③④⑤・⑧充填",
      "重防食塗料の固形分管理", 64,
      "営業窓口", "橋梁・構造物業界経由",
      "DNTと顧客領域が重なる",
      "中堅防食メーカーの補強"),

    # ─── E. 機能・特殊 (6) ────────────────────────────────────
    C("機能・特殊", "オキツモ", "三重県名張市", "非上場",
      "約60億円規模(推定)", "約200人", "耐熱塗料・特殊塗料",
      "◎", "○", "-", "-", "△",
      "三重・名張", "https://www.okitsumo.co.jp/", "本社問合せ",
      "中",
      "耐熱塗料の③分散・⑦検査 (耐熱試験)",
      "シリコーン樹脂取扱、耐熱顔料の分散", 60,
      "技術部門", "プラント・自動車排気系経由",
      "ニッチ専業のため一般化に注意",
      "耐熱塗料のニッチ事例"),
    C("機能・特殊", "シンロイヒ", "東京都品川区", "非上場",
      "約30億円規模(推定)", "約100人", "蛍光塗料・特殊塗料",
      "◎", "△", "-", "-", "△",
      "栃木", "https://www.sinloihi.co.jp/", "本社問合せ",
      "中",
      "蛍光顔料の分散・希釈・充填",
      "蛍光顔料の取扱・退色管理", 58,
      "技術部門", "安全標識・印刷業界経由",
      "蛍光特殊で一般化困難",
      "蛍光特殊事例"),
    C("機能・特殊", "ナトコ", "名古屋市西区", "非上場",
      "約170億円規模(推定)", "約450人", "機能性塗料・特殊コーティング",
      "◎", "○", "△", "-", "○",
      "愛知・茨城", "https://www.natoco.co.jp/", "本社問合せ",
      "中",
      "機能塗料の③分散・⑦検査",
      "機能塗料の品質規格・小ロット運用", 72,
      "技術部門", "電子・自動車・建材業界経由",
      "機能特化のため標準フローと差異あり",
      "機能塗料の補強"),
    C("機能・特殊", "大日精化工業", "東京都中央区", "東証プライム",
      "約1,300億円", "約3,000人(連結)", "顔料・着色剤・UV インキ・特殊塗料",
      "○", "○", "△", "-", "◎",
      "千葉・茨城 ほか", "https://www.daicolorchem.com/", "IR室",
      "高",
      "顔料分散③ + UV 硬化塗料の①〜⑦",
      "顔料分散のビーズ条件・UV ドーズ", 82,
      "技術部門 / IR", "印刷インキ業界経由",
      "インキ事業との切り分けが必要",
      "顔料-塗料の上下流連結の例"),
    C("機能・特殊", "染めQテクノロジィ", "千葉県四街道市", "非上場",
      "約30億円規模(推定)", "約100人", "ナノテク特殊塗料・密着剤",
      "◎", "△", "-", "-", "-",
      "千葉", "https://somayq.com/", "本社問合せ",
      "低〜中",
      "特殊塗料の②分散・⑤調色・⑧充填(エアゾール)",
      "密着剤の薄膜形成、エアゾール充填精度", 60,
      "営業窓口", "DIY・補修業界経由",
      "エアゾール充填が含まれ標準フローと一部異なる",
      "ナノテク特殊事例"),
    C("機能・特殊", "AGCコーテック", "東京都千代田区", "非上場(AGC系)",
      "約60億円規模(推定)", "約200人", "フッ素樹脂塗料・耐候性塗料",
      "◎", "○", "-", "-", "-",
      "千葉・福岡", "https://www.agccoatech.co.jp/", "本社問合せ",
      "中 (AGC連結に含む)",
      "フッ素樹脂塗料の③④⑤",
      "フッ素樹脂取扱・耐候性試験", 70,
      "技術部門 / AGC経由", "建築・インフラ業界経由",
      "フッ素樹脂特化のため一般化に注意",
      "高耐候塗料のニッチ事例"),

    # ─── F. UV・印刷インキ系 (9) ──────────────────────────────
    C("UV・印刷インキ", "十条ケミカル", "東京都北区", "非上場",
      "約100億円規模(推定)", "約300人", "UV インキ・UV 塗料・印刷",
      "△", "△", "-", "-", "◎",
      "東京・埼玉", "https://www.jujochemical.co.jp/", "本社問合せ",
      "中",
      "UV塗料の①遮光下計量・②分散・⑤ろ過・⑦充填",
      "光開始剤の取扱・UV照射ドーズ・遮光容器仕様", 78,
      "技術 / 営業", "印刷業界経由",
      "UVインキとUV塗料の境界に注意",
      "UV硬化の最重要ヒアリング先"),
    C("UV・印刷インキ", "太陽ホールディングス", "埼玉県嵐山町", "東証プライム",
      "約820億円", "約2,200人(連結)", "電子用UVインキ・ソルダーレジスト・UV塗料",
      "△", "△", "-", "-", "◎",
      "埼玉・台湾・中国", "https://www.taiyo-hd.co.jp/", "IR室",
      "高",
      "電子用UV塗料の③分散・⑤ろ過・⑦充填",
      "電子用の微粒子分散・高純度フィルター", 80,
      "IR / 技術部門", "電子業界経由",
      "電子用は機密性高",
      "電子用UVのリファレンス"),
    C("UV・印刷インキ", "帝国インキ製造", "埼玉県川口市", "東証スタンダード",
      "約180億円", "約450人(連結)", "印刷インキ・UV インキ・特殊塗料",
      "△", "△", "-", "-", "◎",
      "川口・千葉", "https://www.teikokuink.com/", "IR室",
      "中〜高",
      "UVインキの③分散・⑦充填",
      "UVインキの分散条件・歩留", 70,
      "IR / 技術", "印刷業界経由",
      "塗料との境界線",
      "UV インキ視点の補強"),
    C("UV・印刷インキ", "T&K TOKA", "埼玉県狭山市", "東証スタンダード",
      "約260億円", "約700人(連結)", "印刷インキ・UV インキ・コーティング",
      "△", "△", "-", "-", "◎",
      "埼玉・大阪", "https://www.tk-toka.co.jp/", "IR室",
      "中〜高",
      "UVコーティングの②分散・⑦充填",
      "UV コーティングのドーズ・密着試験", 70,
      "IR / 技術", "印刷業界経由",
      "塗料との境界が曖昧",
      "UV コーティング系の補強"),
    C("UV・印刷インキ", "東洋インキSCホールディングス", "東京都中央区", "東証プライム",
      "約3,400億円(連結)", "約8,000人(連結)", "印刷インキ・コーティング・ポリマー・色材",
      "○", "○", "△", "-", "◎",
      "埼玉・茨城 ほか", "https://schd.toyoinkgroup.com/", "IR室",
      "高",
      "色材③分散 + UV/水性塗料の②〜⑦",
      "顔料分散のスケール、ポリマー設計", 82,
      "IR / 技術部門", "印刷インキ業界経由",
      "塗料はグループ事業の一部",
      "インキ大手の塗料領域比較対象"),
    C("UV・印刷インキ", "サカタインクス", "大阪市西区", "東証プライム",
      "約1,700億円(連結)", "約4,000人(連結)", "印刷インキ・UV インキ・コーティング",
      "△", "△", "-", "-", "◎",
      "大阪・東京 ほか", "https://www.inx.co.jp/", "IR室",
      "高",
      "UVインキの②分散・⑤ろ過・⑦充填",
      "印刷インキ生産の歩留・廃棄物", 75,
      "IR / 技術", "印刷業界経由",
      "塗料事業は限定的",
      "インキ大手の比較材料"),
    C("UV・印刷インキ", "オリジン", "東京都板橋区", "東証スタンダード",
      "約260億円", "約800人(連結)", "UV 塗料・電着・電子用塗料",
      "○", "○", "△", "○", "◎",
      "東京・神奈川", "https://www.origin.co.jp/", "IR室",
      "中〜高",
      "UV・電着塗料の②③⑤⑦",
      "UV ドーズ管理、電着の電気伝導度", 80,
      "IR / 技術", "電子業界経由",
      "電気事業との切り分け",
      "UV と電着の双方を持つ稀少例"),
    C("UV・印刷インキ", "DICグラフィックス", "東京都中央区", "非上場(DIC子会社)",
      "約2,000億円規模(推定)", "約3,000人", "印刷インキ・UV・特殊インキ",
      "○", "△", "-", "-", "○",
      "千葉・滋賀 ほか", "https://www.dic-graphics.co.jp/", "DIC IR経由",
      "中〜高 (DIC連結に含む)",
      "UV/水性インキの②③⑦",
      "印刷インキの大量生産工程", 72,
      "親会社(DIC) 経由", "印刷業界経由",
      "塗料との境界が曖昧",
      "DIC系のインキ生産規模感"),
    C("UV・印刷インキ", "トウペ", "大阪市平野区", "非上場",
      "約120億円規模(推定)", "約400人", "工業用塗料・UV 塗料・電子部品向け",
      "○", "○", "△", "-", "○",
      "大阪・栃木", "https://www.toupe.co.jp/", "本社問合せ",
      "中",
      "工業用塗料の②③④, UV ⑤⑦",
      "電子向け工業塗料の小ロット運用", 65,
      "営業・技術", "電子・自動車部品業界経由",
      "ニッチ工業塗料は機密性高",
      "工業×UVの中堅事例"),

    # ─── G. DIY・小口 (3) ────────────────────────────────────
    C("DIY・小口", "アサヒペン", "大阪市鶴見区", "非上場",
      "約280億円規模(推定)", "約600人", "DIY・小口・ホームセンター",
      "◎", "◎", "-", "-", "△",
      "大阪・茨城", "https://www.asahipen.jp/", "本社問合せ",
      "中",
      "小容量缶(1L/4L)の⑧充填・⑨保管",
      "小容量充填のスループット・歩留", 60,
      "営業窓口", "ホームセンター経由",
      "プロ向けと差別化必要",
      "DIY 流通の代表"),
    C("DIY・小口", "カンペハピオ", "大阪市東淀川区", "非上場(関ペ系列)",
      "数百億円規模", "数百人", "DIY・ホームセンター・小口",
      "◎", "◎", "-", "-", "-",
      "大阪・関ペ工場併設", "https://www.kanpe.co.jp/", "関ペ広報経由",
      "中",
      "DIY 缶の⑤調色・⑧充填",
      "小容量缶のSKU・色種数・回転", 58,
      "親会社(関ペ) 経由", "ホームセンター業界経由",
      "親会社の方針に従う",
      "DIY 流通の補強"),
    C("DIY・小口", "ニッペホームプロダクツ", "東京都北区", "非上場(日ペ系列)",
      "数百億円規模", "数百人", "DIY・小口・ホームセンター",
      "◎", "◎", "-", "-", "-",
      "東京・関東", "https://www.nippehome.co.jp/", "日ペ広報経由",
      "中",
      "DIY 缶の⑤調色・⑧充填",
      "DIY SKU・小容量充填", 58,
      "親会社(日ペHD) 経由", "ホームセンター業界経由",
      "親会社の方針に従う",
      "DIY 流通の補強(日ペ系)"),

    # ─── H. 日ペグループ子会社 (5) ────────────────────────────
    C("日ペグループ", "日本ペイント・オートモーティブコーティングス", "大阪市北区(日ペHD系)",
      "非上場(子会社)", "日ペHD連結内", "数千人(日ペ系列)", "自動車OEM 塗料 (世界トップクラス)",
      "○", "○", "○", "◎", "△",
      "東京・愛知・大阪", "https://www.nipponpaint-automotive.com/", "日ペHD IR経由",
      "中 (親会社IRに含む)",
      "OEM 向け電着・中塗り・上塗りの①〜⑥",
      "OEM JIT配送、塗着効率、CCM", 92,
      "親会社(日ペHD) 経由", "自動車OEM経由",
      "親会社の方針に従う",
      "自動車OEM の最重要"),
    C("日ペグループ", "日本ペイント・インダストリアルコーティングス", "大阪市北区(日ペHD系)",
      "非上場(子会社)", "日ペHD連結内", "数千人", "工業・電着・特殊塗料",
      "○", "○", "○", "◎", "△",
      "東京・愛知・大阪", "https://www.nipponpaint-industrialcoatings.co.jp/", "日ペHD IR経由",
      "中",
      "工業塗料 ①〜⑨, 電着の特殊工程",
      "電着MEQ・UF・塗着効率、工業ライン", 88,
      "親会社経由", "工業顧客経由",
      "親会社の方針に従う",
      "電着・工業の最重要"),
    C("日ペグループ", "日本ペイントマリン", "兵庫県神戸市(日ペHD系)",
      "非上場(子会社)", "日ペHD連結内", "数千人", "船舶用塗料",
      "◎", "○", "-", "-", "-",
      "兵庫・愛媛", "https://www.nipponpaint-marine.com/", "日ペHD IR経由",
      "中",
      "船舶塗料の③④(防汚機能の分散)",
      "防汚剤・低摩擦塗料の組成", 78,
      "親会社経由", "造船業界経由",
      "中国塗料・関ペマリンと競合",
      "船舶塗料の比較対象"),
    C("日ペグループ", "日本ペイント・サーフケミカルズ", "大阪市北区(日ペHD系)",
      "非上場(子会社)", "日ペHD連結内", "数百人", "表面処理薬剤(塗装前処理)",
      "-", "○", "-", "-", "-",
      "東京・大阪", "https://www.nipponpaint-surf.co.jp/", "日ペHD IR経由",
      "中",
      "塗装前処理薬剤の③配合・⑤希釈",
      "前処理薬剤の濃度管理、リンス回収", 68,
      "親会社経由", "OEM・工業顧客経由",
      "塗料本体ではなく前処理特化",
      "前処理工程の補強"),
    C("日ペグループ", "ニッペ・オートリフィニッシュ", "大阪市北区(日ペHD系)",
      "非上場(子会社)", "日ペHD連結内", "数百人", "自動車補修塗料",
      "◎", "○", "-", "-", "-",
      "東京・大阪", "https://www.nipponpaint-autorefinish.com/", "日ペHD IR経由",
      "中",
      "補修塗料の④⑤調色・⑧充填",
      "小ロット補修の歩留、CCM", 75,
      "親会社経由", "板金塗装業界経由",
      "ロックペイント・イサムと競合",
      "自動車補修の比較対象"),

    # ─── I. 関ペグループ子会社 (1) ────────────────────────────
    C("関ペグループ", "関西ペイントマリン", "大阪市中央区(関ペ系)", "非上場(関ペ系)",
      "関ペ連結内", "数百人", "船舶用塗料",
      "◎", "○", "-", "-", "-",
      "大阪・神奈川", "https://www.kansaipaint-marine.com/", "関ペ広報経由",
      "中",
      "船舶塗料の③④(防汚機能の分散)",
      "防汚剤管理、厚膜塗装", 76,
      "親会社経由", "造船業界経由",
      "中国塗料・日ペマリンと競合",
      "船舶塗料の比較対象"),

    # ─── J. 外資系・日本法人 (8) ──────────────────────────────
    C("外資系・日本法人", "PPGコーティングスジャパン", "大阪市住之江区", "非上場(米PPG子会社)",
      "PPG 連結内", "数千人(日本)", "自動車補修・工業・航空・船舶",
      "◎", "◎", "○", "◎", "○",
      "大阪・愛知 ほか", "https://www.ppgpaint.jp/", "本社問合せ",
      "中 (米親会社IRに開示あり)",
      "OEM・補修・工業塗料の①〜⑦",
      "グローバル標準と日本現場のギャップ", 82,
      "技術部門 / 営業", "OEM・補修・工業経由",
      "親会社の方針に従う・グローバル機密",
      "外資系の代表"),
    C("外資系・日本法人", "アクゾノーベル・ジャパン", "東京都港区", "非上場(蘭AkzoNobel子会社)",
      "AkzoNobel連結内", "数百人(日本)", "建築・船舶・工業・粉体",
      "◎", "◎", "◎", "○", "○",
      "東京・大阪 ほか", "https://www.akzonobel.com/ja/", "本社問合せ",
      "中",
      "粉体・水性建築塗料の②〜⑦",
      "粉体グローバル標準の歩留・能力", 78,
      "技術部門 / 営業", "業界団体・OEM経由",
      "親会社方針・グローバル機密",
      "粉体の世界トップの日本拠点"),
    C("外資系・日本法人", "シャーウィン・ウィリアムズ・ジャパン", "東京都", "非上場(米SW子会社)",
      "SW連結内", "数百人(日本)", "建築・工業・自動車補修",
      "○", "◎", "○", "○", "△",
      "(要確認)", "https://www.sherwin-williams.com/", "本社問合せ",
      "低〜中 (要確認)",
      "建築塗料 (米国基準) と日本基準の比較",
      "米国基準塗料の日本現場運用", 70,
      "本社経由", "(限定的)",
      "事業規模は要確認",
      "外資建築塗料の参考"),
    C("外資系・日本法人", "アクサルタコーティングシステムズ・ジャパン", "横浜市", "非上場(米Axalta子会社)",
      "Axalta連結内", "数百人(日本)", "自動車OEM・補修・工業",
      "◎", "○", "○", "◎", "△",
      "横浜・大阪", "https://www.axalta.com/jp/", "本社問合せ",
      "中",
      "OEM・補修・電着の①〜⑥",
      "電着のグローバル標準、OEM JIT", 78,
      "技術部門 / 営業", "OEM・補修経由",
      "親会社方針・グローバル機密",
      "電着・OEM の外資比較対象"),
    C("外資系・日本法人", "BASFジャパン コーティングス事業", "東京都中央区", "非上場(独BASF子会社)",
      "BASF Coatings連結内", "数百人(日本)", "自動車OEM・補修・特殊",
      "◎", "○", "△", "○", "△",
      "(要確認・事業再編動向あり)", "https://www.basf-coatings.com/", "本社問合せ",
      "低〜中 (事業再編動向で要確認)",
      "OEM 塗料の①〜⑥",
      "OEM 塗料の高機能化動向", 72,
      "本社経由", "自動車OEM経由",
      "コーティング事業の売却動向 (2024〜) に注意",
      "事業再編動向の継続確認が必要"),
    C("外資系・日本法人", "ヘンペル・ジャパン", "横浜市", "非上場(デンマークHempel子会社)",
      "Hempel連結内", "数十人(日本)", "船舶・防食・コンテナ用塗料",
      "◎", "○", "-", "-", "-",
      "横浜", "https://www.hempel.com/ja-jp/", "本社問合せ",
      "中",
      "船舶塗料の③④",
      "船舶塗料グローバル標準", 65,
      "本社経由", "造船・港湾経由",
      "本社方針優先",
      "船舶塗料の欧州比較"),
    C("外資系・日本法人", "ヨトゥン・ジャパン", "横浜市", "非上場(ノルウェーJotun子会社)",
      "Jotun連結内", "数十人(日本)", "船舶・防食・建築用塗料",
      "◎", "○", "-", "-", "-",
      "横浜", "https://www.jotun.com/jp/", "本社問合せ",
      "中",
      "船舶塗料の③④, 建築水性②⑤",
      "船舶・建築双方のグローバル標準", 65,
      "本社経由", "造船・建築業界経由",
      "本社方針優先",
      "船舶・建築の欧州比較"),
    C("外資系・日本法人", "ベンジャミンムーア・ジャパン", "東京都", "非上場(米BM子会社)",
      "BM連結内", "数十人(日本)", "建築用高級水性塗料",
      "△", "◎", "-", "-", "-",
      "(輸入主体)", "https://benjaminmoore.jp/", "本社問合せ",
      "低",
      "建築水性塗料の調色・小口販売",
      "米国の調色機普及・小口運用", 55,
      "輸入代理店経由", "建材専門店経由",
      "輸入主体で国内製造は限定的",
      "輸入塗料の参考"),

    # ─── K. 追加で計50社 (機能・特殊系を 1社追加) ─────────────
    C("機能・特殊", "ターナー色彩", "大阪府吹田市", "非上場",
      "約60億円規模(推定)", "約200人", "アクリル絵具・アート用塗料・建築仕上げ",
      "◎", "◎", "-", "-", "-",
      "大阪・吹田", "https://www.turner.co.jp/", "本社問合せ",
      "中",
      "アクリル系塗料の③④⑤⑧(チューブ・ボトル充填)",
      "小ロット多品種(色数)管理、特殊容器充填", 60,
      "営業窓口", "画材・建材販売店経由",
      "アート用と建築用の両方を持つ特殊性",
      "多品種小ロットの極端例"),
]


def make_cover_sheet(wb):
    ws = wb.create_sheet("表紙_運用方針", 0)
    setup_columns(ws, [3, 30, 50, 30])

    ws.merge_cells("B2:D2")
    cell(ws, 2, 2, "塗料メーカー ヒアリング候補リスト (網羅型・50社)",
         font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 30

    ws.merge_cells("B3:D3")
    cell(ws, 3, 2,
         "塗料製造フロー & IN/OUT (別Excel) の記述を、塗料メーカーへのヒアリングで検証するための候補企業リスト (業界団体・装置メーカーは含めず塗料メーカーのみ)",
         font=Font(name="Yu Gothic", size=10, italic=True, color="555555"),
         align=ALIGN_CENTER, border=None)

    cell(ws, 5, 2, "確定方針", font=FONT_H, fill=FILL_HDR1, align=ALIGN_CENTER)
    ws.merge_cells("B5:D5")
    rows = [
        ("対象", "塗料メーカーのみ 50社", "業界団体・装置メーカーは除外"),
        ("アプローチ", "A. 網羅型", "11カテゴリで 5製品タイプ × 規模 × セグメントを網羅"),
        ("ヒアリング立場", "業界関係者", "取引先・関連事業者として既存ネットワーク活用"),
        ("成果物", "Plan-1: リスト Excel のみ", "ヒアリング項目は別途自社で設計"),
        ("検証深さ", "定性 + 定量", "フロー/項目正誤 + kWh/t・歩留・廃棄量校正"),
    ]
    for i, (k, v, note) in enumerate(rows, start=6):
        cell(ws, i, 2, k, font=FONT_H, align=ALIGN_LEFT)
        cell(ws, i, 3, v, font=FONT_BODY, align=ALIGN_LEFT)
        cell(ws, i, 4, note, font=FONT_BODY, align=ALIGN_WRAP)
        ws.row_dimensions[i].height = 28

    cell(ws, 12, 2, "シート構成", font=FONT_H, fill=FILL_HDR2, align=ALIGN_CENTER)
    ws.merge_cells("B12:D12")
    rows = [
        ("①", "ヒアリング先 候補リスト (50社)",
         "11カテゴリ × 26列。基本情報・規模・5製品タイプ適合・工場・連絡・ヒアリング推奨工程・優先度スコア・アプローチ・既存関係 etc."),
        ("②", "優先度スコアリング基準",
         "100点満点の配点設計。カバレッジ・規模・公開度・アクセス・専門特化価値"),
        ("③", "カテゴリ別カバレッジ確認",
         "5製品タイプ × カテゴリ別の ◎/○ 保有社数を集計し、漏れ・偏りを確認"),
    ]
    for i, (no, name, desc) in enumerate(rows, start=13):
        cell(ws, i, 2, no, font=FONT_H, align=ALIGN_CENTER)
        cell(ws, i, 3, name, font=FONT_BODY, align=ALIGN_LEFT)
        cell(ws, i, 4, desc, font=FONT_BODY, align=ALIGN_WRAP)
        ws.row_dimensions[i].height = 38

    cell(ws, 17, 2, "リスト利用時の重要注意事項", font=FONT_H, fill=FILL_HDR3, align=ALIGN_CENTER)
    ws.merge_cells("B17:D17")
    notes = [
        "・本リストの企業情報 (売上・従業員数・セグメント・適合製品タイプ) は 公開情報ベースの初期案 (推定値を含む)。",
        "・実ヒアリング前に貴社の取引履歴・関係性で 「既存関係」「優先度スコア」「アプローチルート」 列を必ず上書き校正してください。",
        "・「公開情報充実度」列に \"(要確認)\" がある企業は、基本情報の出典再確認が必要です。",
        "・「過去契約有無」列は kintone 取引明細 (2026/05/11 抽出) との突合結果。○ 印=過去契約あり。",
        "  突合は親会社のみ厳密マッチ。日ペHDの取引があっても子会社5社には自動付与しない (逆も同様)。",
        f"  ○ 印が付いた企業: 8社/{50}社 (関西ペイント, ロックペイント, ナトコ, シンロイヒ, 大日精化工業, サカタインクス, トウペ, 日本ペイント・インダストリアルコーティングス)",
        "・優先度スコアは 100 点満点。シート②に配点基準を記載。",
        "・売上・従業員等は推定含む。最新の有報・IR資料・帝国データバンク等で適宜上書きを推奨。",
        "・◎=主力 / ○=取扱 / △=限定 / -=無し  (5製品タイプ別)。",
        "・「ヒアリング推奨工程」は別 Excel 「塗料メーカー_製造フローINOUT.xlsx」の工程番号と連動。",
        "・日ペグループ・関ペグループ子会社は、独立にアプローチせず親会社経由を推奨。",
        "・外資系日本法人は本社方針・グローバル機密の制約が強い。日本拠点での製造規模を要確認。",
    ]
    for i, n in enumerate(notes, start=18):
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        cell(ws, i, 2, n, font=FONT_BODY, align=ALIGN_LEFT)
        ws.row_dimensions[i].height = 22


def make_main_list_sheet(wb):
    ws = wb.create_sheet("①候補リスト")
    widths = [
        4,   # No.
        14,  # カテゴリ
        32,  # 企業名
        12,  # 過去契約有無 (新規)
        20,  # 本社所在地
        18,  # 上場区分
        20,  # 売上規模
        16,  # 従業員数
        26,  # 主要セグメント
        8, 8, 8, 8, 8,   # 製品タイプ適合 5列
        24,  # 主要工場
        32,  # Web
        22,  # IR/問合せ
        18,  # 公開情報充実度
        30,  # ヒアリング推奨工程
        32,  # 確認したい定量項目
        10,  # 優先度スコア
        24,  # 推奨アプローチ
        22,  # 既存関係
        30,  # 留意点
        14,  # 進捗ステータス
        20,  # 次アクション
        24,  # メモ
    ]
    setup_columns(ws, widths)

    last_col = len(HEADERS)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    cell(ws, 1, 1, f"① 塗料メーカー ヒアリング先 候補リスト (網羅型 {len(COMPANIES)}社)",
         font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    cell(ws, 2, 1,
         "◎=主力ライン / ○=取扱あり / △=限定的 / -=取扱なし  ｜  優先度スコアの基準はシート②を参照  ｜  数値・規模は推定含む・要校正",
         font=Font(name="Yu Gothic", size=9, italic=True, color="555555"),
         align=ALIGN_CENTER, border=None)
    ws.row_dimensions[2].height = 18

    for i, h in enumerate(HEADERS, start=1):
        cell(ws, 4, i, h, font=FONT_H, fill=FILL_HDR1, align=ALIGN_CENTER)
    ws.row_dimensions[4].height = 36

    for ci, comp in enumerate(COMPANIES, start=5):
        fill = CAT_FILL.get(comp["カテゴリ"], None)
        cell(ws, ci, 1, ci - 4, font=FONT_BODY, align=ALIGN_CENTER, fill=fill)
        cell(ws, ci, 2, comp["カテゴリ"], font=FONT_BODY, align=ALIGN_CENTER, fill=fill)
        cell(ws, ci, 3, comp["企業名"],
             font=Font(name="Yu Gothic", size=10, bold=True), align=ALIGN_LEFT)
        # 4. 過去契約有無 (新規列)
        name = comp["企業名"]
        pc_n = PAST_CONTRACT.get(name, 0)
        pc_cell = cell(ws, ci, 4, "○" if pc_n > 0 else "",
                       font=Font(name="Yu Gothic", size=14, bold=True, color="C00000"),
                       align=ALIGN_CENTER)
        if pc_n > 0:
            pc_cell.fill = PatternFill("solid", fgColor="C6EFCE")
        cell(ws, ci, 5, comp["本社所在地"], font=FONT_BODY, align=ALIGN_LEFT)
        cell(ws, ci, 6, comp["上場区分"], font=FONT_BODY, align=ALIGN_CENTER)
        cell(ws, ci, 7, comp["売上規模(連結)"], font=FONT_BODY, align=ALIGN_LEFT)
        cell(ws, ci, 8, comp["従業員数(概数)"], font=FONT_BODY, align=ALIGN_LEFT)
        cell(ws, ci, 9, comp["主要セグメント"], font=FONT_BODY, align=ALIGN_WRAP)
        for k_off, key in enumerate(["溶剤系", "水性", "粉体", "電着", "UV硬化"]):
            v = comp[key]
            fcell = cell(ws, ci, 10 + k_off, v, font=FONT_H, align=ALIGN_CENTER)
            if v == "◎":
                fcell.fill = PatternFill("solid", fgColor="C6E0B4")
            elif v == "○":
                fcell.fill = PatternFill("solid", fgColor="DDEBF7")
            elif v == "△":
                fcell.fill = PatternFill("solid", fgColor="FFF2CC")
        cell(ws, ci, 15, comp["主要工場(代表)"], font=FONT_BODY, align=ALIGN_WRAP)
        c = cell(ws, ci, 16, comp["Web"], font=FONT_BODY, align=ALIGN_LEFT)
        if comp["Web"].startswith("http"):
            c.hyperlink = comp["Web"]
            c.font = Font(name="Yu Gothic", size=9, color="0563C1", underline="single")
        cell(ws, ci, 17, comp["IR/問合せ窓口"], font=FONT_BODY, align=ALIGN_LEFT)
        # 公開情報充実度: "(要確認)" を含む場合は強調
        info_v = comp["公開情報充実度"]
        info_cell = cell(ws, ci, 18, info_v, font=FONT_BODY, align=ALIGN_CENTER)
        if "要確認" in info_v:
            info_cell.fill = PatternFill("solid", fgColor="FFC7CE")
            info_cell.font = Font(name="Yu Gothic", size=9, bold=True, color="9C0006")
        cell(ws, ci, 19, comp["ヒアリング推奨工程"], font=FONT_BODY, align=ALIGN_WRAP)
        cell(ws, ci, 20, comp["確認したい定量項目"], font=FONT_BODY, align=ALIGN_WRAP)
        score = comp["優先度スコア"]
        sc = cell(ws, ci, 21, score, font=FONT_NUM, align=ALIGN_CENTER)
        if score >= 85:
            sc.fill = PatternFill("solid", fgColor="C6EFCE")
        elif score >= 75:
            sc.fill = PatternFill("solid", fgColor="FFEB9C")
        elif score >= 65:
            sc.fill = PatternFill("solid", fgColor="FFC7CE")
        else:
            sc.fill = PatternFill("solid", fgColor="EDEDED")
        cell(ws, ci, 22, comp["推奨アプローチルート"], font=FONT_BODY, align=ALIGN_WRAP)
        cell(ws, ci, 23, comp["既存関係(想定)"], font=FONT_BODY, align=ALIGN_WRAP)
        cell(ws, ci, 24, comp["留意点・ハードル"], font=FONT_BODY, align=ALIGN_WRAP)
        cell(ws, ci, 25, comp["進捗ステータス"], font=FONT_BODY, align=ALIGN_CENTER)
        cell(ws, ci, 26, comp["次アクション"], font=FONT_BODY, align=ALIGN_LEFT)
        cell(ws, ci, 27, comp["メモ"], font=FONT_BODY, align=ALIGN_WRAP)
        ws.row_dimensions[ci].height = 68

    ws.auto_filter.ref = f"A4:{get_column_letter(last_col)}{4 + len(COMPANIES)}"
    ws.freeze_panes = "E5"

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def make_scoring_sheet(wb):
    ws = wb.create_sheet("②優先度スコアリング基準")
    setup_columns(ws, [3, 26, 10, 60])

    ws.merge_cells("B2:D2")
    cell(ws, 2, 2, "② 優先度スコアリング基準 (100点満点)",
         font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 28

    headers = ["評価軸", "配点", "評価方法 / 目安"]
    for i, h in enumerate(headers, start=2):
        cell(ws, 4, i, h, font=FONT_H, fill=FILL_HDR1, align=ALIGN_CENTER)

    rows = [
        ("カバレッジ (製品タイプ網羅性)", 30,
         "5製品タイプのうち◎または○保有数。5=30点 / 4=24 / 3=18 / 2=12 / 1=6"),
        ("規模・代表性", 20,
         "売上順位・国内シェア。大手2社=20点 / 中堅=15 / 中小=10 / ニッチ専業=8"),
        ("公開度 (一次情報の取りやすさ)", 20,
         "統合報告書・環境/サステナ報告書・IR充実度。高=20 / 中=12 / 低=6 / 要確認=減点"),
        ("アクセス容易性", 15,
         "既存ネットワーク・業界団体経由・地理的近接性。高=15 / 中=10 / 低=5"),
        ("専門特化価値", 15,
         "他で代替不可能な情報源(船舶・電着・UV・粉体・耐熱等)。高=15 / 中=10 / 低=5"),
    ]
    for i, (a, p, m) in enumerate(rows, start=5):
        cell(ws, i, 2, a, font=FONT_H, align=ALIGN_LEFT)
        cell(ws, i, 3, p, font=FONT_NUM, align=ALIGN_CENTER)
        cell(ws, i, 4, m, font=FONT_BODY, align=ALIGN_WRAP)
        ws.row_dimensions[i].height = 35

    cell(ws, 10, 2, "合計", font=FONT_H, fill=FILL_HDR2, align=ALIGN_CENTER)
    cell(ws, 10, 3, 100, font=Font(name="Yu Gothic", size=11, bold=True),
         fill=FILL_HDR2, align=ALIGN_CENTER)
    cell(ws, 10, 4, "全社のスコアを合算して並び替え、上位から打診計画を作成",
         font=FONT_BODY, fill=FILL_HDR2, align=ALIGN_LEFT)

    cell(ws, 12, 2, "スコア帯と意味", font=FONT_H, fill=FILL_HDR3, align=ALIGN_CENTER)
    ws.merge_cells("B12:D12")
    ranks = [
        ("85〜100点", "S: 最優先 (まず着手)",       "C6EFCE", "業界平均/標準化に最も影響大"),
        ("75〜84点",  "A: 重点候補",               "FFEB9C", "代表性ある中堅・特化メーカー"),
        ("65〜74点",  "B: 補強候補",               "FFC7CE", "他で取れない補強情報を期待"),
        ("〜64点",    "C: 補完 (必要に応じて)",    "EDEDED", "他社で代替可能ならスキップも可"),
    ]
    for i, (band, label, color, note) in enumerate(ranks, start=13):
        cell(ws, i, 2, band, font=FONT_H, align=ALIGN_CENTER,
             fill=PatternFill("solid", fgColor=color))
        cell(ws, i, 3, label, font=FONT_BODY, align=ALIGN_LEFT,
             fill=PatternFill("solid", fgColor=color))
        cell(ws, i, 4, note, font=FONT_BODY, align=ALIGN_LEFT,
             fill=PatternFill("solid", fgColor=color))
        ws.row_dimensions[i].height = 25


def make_coverage_sheet(wb):
    ws = wb.create_sheet("③カバレッジ確認")
    setup_columns(ws, [3, 24, 12, 12, 12, 12, 12, 8, 30])

    ws.merge_cells("B2:I2")
    cell(ws, 2, 2, "③ 5製品タイプ × カテゴリ別 カバレッジ確認",
         font=FONT_TITLE, fill=FILL_TITLE, align=ALIGN_CENTER, border=BORDER_BOX)
    ws.row_dimensions[2].height = 28

    ws.merge_cells("B3:I3")
    cell(ws, 3, 2,
         "各製品タイプの ◎ または ○ 保有社数を集計。漏れ・過剰がないか確認するための表。",
         font=Font(name="Yu Gothic", size=9, italic=True, color="555555"),
         align=ALIGN_CENTER, border=None)
    ws.row_dimensions[3].height = 18

    headers = ["カテゴリ", "溶剤系", "水性", "粉体", "電着", "UV硬化", "社数", "備考"]
    for i, h in enumerate(headers, start=2):
        cell(ws, 5, i, h, font=FONT_H, fill=FILL_HDR1, align=ALIGN_CENTER)

    # カテゴリ順をデータ通りに
    order = []
    for c in COMPANIES:
        if c["カテゴリ"] not in order:
            order.append(c["カテゴリ"])

    for ci, cat in enumerate(order, start=6):
        cat_comps = [c for c in COMPANIES if c["カテゴリ"] == cat]
        cell(ws, ci, 2, cat,
             font=FONT_H, align=ALIGN_LEFT, fill=CAT_FILL.get(cat, None))
        for col_off, key in enumerate(["溶剤系", "水性", "粉体", "電着", "UV硬化"]):
            cm = sum(1 for c in cat_comps if c[key] == "◎")
            cs = sum(1 for c in cat_comps if c[key] == "○")
            v = f"◎{cm} / ○{cs}" if (cm + cs) > 0 else "-"
            cell(ws, ci, 3 + col_off, v, font=FONT_BODY, align=ALIGN_CENTER)
        cell(ws, ci, 8, len(cat_comps), font=FONT_NUM, align=ALIGN_CENTER)
        cell(ws, ci, 9, "", font=FONT_BODY, align=ALIGN_LEFT)
        ws.row_dimensions[ci].height = 26

    total_row = 6 + len(order)
    cell(ws, total_row, 2, "全社合計",
         font=FONT_H, align=ALIGN_LEFT, fill=FILL_HDR2)
    for col_off, key in enumerate(["溶剤系", "水性", "粉体", "電着", "UV硬化"]):
        cm = sum(1 for c in COMPANIES if c[key] == "◎")
        cs = sum(1 for c in COMPANIES if c[key] == "○")
        cell(ws, total_row, 3 + col_off, f"◎{cm} / ○{cs}",
             font=FONT_NUM, align=ALIGN_CENTER, fill=FILL_HDR2)
    cell(ws, total_row, 8, len(COMPANIES),
         font=Font(name="Yu Gothic", size=11, bold=True, color="C00000"),
         align=ALIGN_CENTER, fill=FILL_HDR2)
    cell(ws, total_row, 9,
         "◎○ 合計が 5 社未満の製品タイプは要補強",
         font=FONT_BODY, align=ALIGN_LEFT, fill=FILL_HDR2)
    ws.row_dimensions[total_row].height = 26


def main():
    wb = Workbook()
    wb.remove(wb.active)
    make_cover_sheet(wb)
    make_main_list_sheet(wb)
    make_scoring_sheet(wb)
    make_coverage_sheet(wb)
    wb.save(OUTPUT)
    print(f"OK: {OUTPUT}  ({len(COMPANIES)}社収録)")


if __name__ == "__main__":
    main()
