"""Create public, data-free promotional graphics for Seqara PPT Personal.

All chart marks, photo blocks, and slide frames are illustrative. They are not
screenshots of a generated presentation or depictions of a real institution.
"""

from __future__ import annotations

from pathlib import Path

import cairosvg


OUT = Path(__file__).resolve().parent


def svg_shell(width: int, height: int, body: str) -> str:
    return f'''<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}" role="img">
<defs>
  <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#071625"/><stop offset="0.56" stop-color="#10243B"/><stop offset="1" stop-color="#071522"/></linearGradient>
  <linearGradient id="panel" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#193450"/><stop offset="1" stop-color="#0D2136"/></linearGradient>
  <linearGradient id="photo" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#93A6B7"/><stop offset="1" stop-color="#314B63"/></linearGradient>
  <linearGradient id="amber" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#FFE09A"/><stop offset="1" stop-color="#E3A34B"/></linearGradient>
  <filter id="glow"><feGaussianBlur stdDeviation="22"/></filter>
</defs>
<rect width="{width}" height="{height}" fill="url(#bg)"/>
<circle cx="{int(width*.83)}" cy="{int(height*.12)}" r="{int(height*.28)}" fill="#57D4DF" opacity=".08" filter="url(#glow)"/>
<circle cx="{int(width*.26)}" cy="{int(height*.95)}" r="{int(height*.23)}" fill="#F4AD58" opacity=".07" filter="url(#glow)"/>
<g opacity=".2" stroke="#7393A9" stroke-width="1"><path d="M0 136 H{width}"/><path d="M0 {height-102} H{width}"/></g>
{body}
</svg>'''


def txt(x: int, y: int, s: str, size: int, *, fill: str = '#F2F5F7', weight: int = 400, tracking: int = 0) -> str:
    import html

    return (f'<text x="{x}" y="{y}" fill="{fill}" font-family="Noto Sans SC, Microsoft YaHei, sans-serif" '
            f'font-size="{size}" font-weight="{weight}" letter-spacing="{tracking}">{html.escape(s)}</text>')


def chart_card(x: int, y: int, scale: float = 1) -> str:
    return f'''<g transform="translate({x} {y}) scale({scale})">
<rect width="510" height="300" rx="24" fill="#F8F9F9"/>
<rect width="510" height="55" rx="24" fill="#EDF1F3"/><rect y="30" width="510" height="25" fill="#EDF1F3"/>
<circle cx="30" cy="28" r="7" fill="#E8AA52"/><circle cx="51" cy="28" r="7" fill="#A7B8C6"/><circle cx="72" cy="28" r="7" fill="#A7B8C6"/>
<rect x="30" y="79" width="210" height="10" rx="5" fill="#12283D"/>
<rect x="30" y="102" width="135" height="7" rx="3" fill="#9FAEB9"/>
<path d="M47 247H475 M47 197H475 M47 147H475" stroke="#DFE6E9" stroke-width="2"/>
<path d="M64 221 L154 193 L247 204 L340 152 L443 132" fill="none" stroke="#EEA951" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>
<circle cx="64" cy="221" r="7" fill="#EEA951"/><circle cx="154" cy="193" r="7" fill="#EEA951"/><circle cx="247" cy="204" r="7" fill="#EEA951"/><circle cx="340" cy="152" r="7" fill="#EEA951"/><circle cx="443" cy="132" r="7" fill="#EEA951"/>
<rect x="60" y="259" width="40" height="5" rx="2" fill="#A7B8C6"/><rect x="159" y="259" width="40" height="5" rx="2" fill="#A7B8C6"/><rect x="258" y="259" width="40" height="5" rx="2" fill="#A7B8C6"/><rect x="357" y="259" width="40" height="5" rx="2" fill="#A7B8C6"/>
</g>'''


def narrative_card(x: int, y: int, scale: float = 1) -> str:
    return f'''<g transform="translate({x} {y}) scale({scale})">
<rect width="370" height="250" rx="22" fill="url(#panel)" stroke="#4C6983" stroke-width="2"/>
<rect x="33" y="38" width="88" height="7" rx="3" fill="#E9AB5D"/>
<rect x="33" y="72" width="255" height="16" rx="5" fill="#ECF2F5"/>
<rect x="33" y="102" width="194" height="8" rx="3" fill="#7F9CB4"/>
<path d="M41 179H330" stroke="#53738C" stroke-width="4"/>
<circle cx="62" cy="179" r="12" fill="#E9AB5D"/><circle cx="179" cy="179" r="12" fill="#7ECED7"/><circle cx="310" cy="179" r="12" fill="#E9AB5D"/>
<path d="M179 178l36 -35" fill="none" stroke="#7ECED7" stroke-width="3"/>
<rect x="35" y="211" width="76" height="6" rx="3" fill="#7F9CB4"/><rect x="151" y="211" width="76" height="6" rx="3" fill="#7F9CB4"/><rect x="270" y="211" width="58" height="6" rx="3" fill="#7F9CB4"/>
</g>'''


def image_card(x: int, y: int, scale: float = 1) -> str:
    return f'''<g transform="translate({x} {y}) scale({scale})">
<rect width="310" height="240" rx="21" fill="#F6F7F7"/>
<rect x="22" y="25" width="266" height="24" rx="4" fill="#E7A856"/>
<rect x="22" y="70" width="119" height="124" rx="8" fill="url(#photo)"/>
<rect x="153" y="70" width="135" height="56" rx="8" fill="#BAC5CC"/>
<rect x="153" y="138" width="135" height="56" rx="8" fill="#516B7D"/>
<path d="M37 183l34 -45 25 25 25 -35" fill="none" stroke="#DCE7E9" stroke-width="5" opacity=".7"/>
<rect x="22" y="210" width="144" height="6" rx="3" fill="#60798C"/>
</g>'''


def hero() -> str:
    w, h = 1600, 900
    parts = [
        '<rect x="78" y="74" width="207" height="34" rx="17" fill="#1E3C53" stroke="#48728B"/>',
        txt(103, 99, 'PERSONAL SKILL', 17, fill='#B8D8E1', weight=700, tracking=2),
        txt(78, 207, '让口述稿有证据、', 64, weight=700),
        txt(78, 294, '有节奏、有画面。', 64, weight=700),
        txt(82, 354, 'Seqara PPT Personal', 34, fill='#EAB268', weight=600),
        txt(82, 407, '原生可编辑 · 来源可追溯 · 逐页可核验', 22, fill='#B8C8D4'),
        '<path d="M86 490H684" stroke="#49667B" stroke-width="2"/>',
        txt(83, 546, '01  数据与图表', 25, weight=600),
        txt(83, 585, '指标绑定来源，五类原生图表。', 20, fill='#AFC1CE'),
        txt(83, 645, '02  叙事与转场', 25, weight=600),
        txt(83, 684, '按讲述意群决定页面节奏。', 20, fill='#AFC1CE'),
        txt(83, 744, '03  项目图片', 25, weight=600),
        txt(83, 783, '提取、匹配、排版，歧义待确认。', 20, fill='#AFC1CE'),
        '<g opacity=".36"><rect x="854" y="235" width="625" height="370" rx="30" fill="#284B67" transform="rotate(5 1165 420)"/></g>',
        chart_card(805, 180, 1.28),
        narrative_card(860, 555, .91),
        image_card(1225, 515, .87),
        '<rect x="79" y="829" width="579" height="34" rx="17" fill="#19374C"/>',
        txt(101, 852, '功能示意 · 非真实项目或实际生成截图', 17, fill='#AFC1CE'),
    ]
    return svg_shell(w, h, '\n'.join(parts))


def social() -> str:
    w, h = 1200, 630
    parts = [
        '<rect x="56" y="53" width="173" height="29" rx="14" fill="#24455D"/>',
        txt(78, 75, 'PERSONAL SKILL', 14, fill='#B8D8E1', weight=700, tracking=2),
        txt(56, 176, '口述稿 → 可编辑 PPT', 55, weight=700),
        txt(58, 233, '有证据的数据 · 有节奏的叙事 · 有来源的图片', 23, fill='#B9CAD4'),
        '<path d="M58 275H600" stroke="#5B7388" stroke-width="2"/>',
        txt(59, 336, 'Seqara PPT Personal', 35, fill='#F0B46C', weight=700),
        '<rect x="58" y="374" width="488" height="119" rx="18" fill="#173149" stroke="#31536E"/>',
        txt(84, 417, '原生对象复用  ·  五类可编辑图表', 21, weight=500),
        txt(84, 458, '项目图归属审核  ·  PowerPoint 渲染检查', 20, fill='#B5C6D2'),
        chart_card(662, 141, .90),
        narrative_card(771, 423, .48),
        image_card(1008, 405, .43),
        txt(59, 595, '功能示意 · 非真实项目截图', 15, fill='#809BAE'),
    ]
    return svg_shell(w, h, '\n'.join(parts))


def feature_square() -> str:
    w, h = 1080, 1080
    parts = [
        '<rect x="62" y="60" width="211" height="33" rx="16" fill="#24455D"/>',
        txt(85, 84, 'PERSONAL SKILL', 16, fill='#B8D8E1', weight=700, tracking=2),
        txt(62, 181, '一份讲稿，三层表达。', 56, weight=700),
        txt(65, 233, '用原生对象把内容变成可复核的汇报。', 23, fill='#AFC5D1'),
        '<rect x="60" y="282" width="960" height="206" rx="26" fill="#173149" stroke="#3B5C74"/>',
        txt(90, 336, '01  数据有依据', 31, fill='#F0B46C', weight=700),
        txt(90, 379, '表格单元格 → 指标口径 → 可编辑图表', 22),
        '<path d="M676 426H963" stroke="#66859A" stroke-width="3"/>',
        '<path d="M693 412l59 -43 62 30 65 -74 65 22" fill="none" stroke="#E8A958" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>',
        '<rect x="60" y="512" width="960" height="206" rx="26" fill="#173149" stroke="#3B5C74"/>',
        txt(90, 568, '02  节奏有逻辑', 31, fill='#F0B46C', weight=700),
        txt(90, 611, '讲述意群 → 桥接句 → 必要的转场页', 22),
        '<path d="M665 647H959" stroke="#66859A" stroke-width="5"/>',
        '<circle cx="698" cy="647" r="13" fill="#E9AB5D"/><circle cx="813" cy="647" r="13" fill="#7ECED7"/><circle cx="928" cy="647" r="13" fill="#E9AB5D"/>',
        '<rect x="60" y="742" width="960" height="206" rx="26" fill="#173149" stroke="#3B5C74"/>',
        txt(90, 798, '03  图片有出处', 31, fill='#F0B46C', weight=700),
        txt(90, 841, 'Word 提取 → 项目归属 → 裁切复核', 22),
        '<rect x="684" y="784" width="100" height="109" rx="9" fill="#7991A2"/><rect x="795" y="784" width="163" height="50" rx="9" fill="#B7C8D0"/><rect x="795" y="844" width="163" height="49" rx="9" fill="#49677B"/>',
        txt(63, 1029, '功能示意 · 非真实项目截图', 16, fill='#809BAE'),
    ]
    return svg_shell(w, h, '\n'.join(parts))


def main() -> None:
    for name, source in [('hero', hero()), ('social-cover', social()), ('feature-square', feature_square())]:
        svg_path = OUT / f'{name}.svg'
        png_path = OUT / f'{name}.png'
        svg_path.write_text(source, encoding='utf-8')
        cairosvg.svg2png(bytestring=source.encode('utf-8'), write_to=str(png_path))
        print(svg_path, png_path)


if __name__ == '__main__':
    main()
