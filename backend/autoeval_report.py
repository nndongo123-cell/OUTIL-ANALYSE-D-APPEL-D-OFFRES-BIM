from __future__ import annotations

import base64
import html
import io
from pathlib import Path
from typing import Any, Dict, Mapping

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image


def _clean(value: Any) -> str:
    import re
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def _msg(messages: Mapping[str, str] | None, key: str, **ctx: Any) -> str:
    template = _clean((messages or {}).get(key, ''))
    if not template:
        return ''
    class Safe(dict):
        def __missing__(self, name):
            return ''
    try:
        return _clean(template.format_map(Safe({k: _clean(v) for k, v in ctx.items()})))
    except Exception:
        return template


def _axis_label(axes: Mapping[str, Any] | None, axis: str, code: str) -> str:
    return _clean((((axes or {}).get(axis) or {}).get(code) or {}).get('label')) or _clean(code)


def _radar_png(scores: Mapping[str, Mapping[str, Any]], title: str, params: Mapping[str, Any] | None = None) -> bytes:
    params = params or {}
    items = [(bid, data) for bid, data in scores.items() if isinstance(data.get('pct'), (int, float))]
    if not items:
        return b''
    try:
        scale_max = float(params.get('radar_score_max'))
    except (TypeError, ValueError):
        return b''
    if scale_max <= 0:
        return b''
    raw_ticks = params.get('radar_graduations') or []
    if isinstance(raw_ticks, str):
        raw_ticks = [v.strip() for v in raw_ticks.split(';') if v.strip()]
    ticks = []
    for raw in raw_ticks:
        try:
            value = float(raw)
        except (TypeError, ValueError):
            continue
        if 0 < value <= scale_max:
            ticks.append(value)
    labels = [bid for bid, _ in items]
    values = [max(0.0, min(scale_max, float(data.get('pct', 0)))) / scale_max for _, data in items]
    n = len(items)
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]
    values += values[:1]
    fig, ax = plt.subplots(figsize=(7.2, 5.8), subplot_kw={'polar': True})
    # Palette identique au radar historique du dashboard : présentation uniquement.
    ax.plot(angles, values, linewidth=2.5, color='#A1003D')
    ax.fill(angles, values, alpha=0.14, color='#A1003D')
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=8, color='#17232B', fontweight='bold')
    ax.set_ylim(0, 1)
    ax.set_yticks([v / scale_max for v in ticks])
    ax.set_yticklabels([f'{v:g} %' for v in ticks], fontsize=7, color='#667985')
    ax.grid(color='#D9E3E9', linewidth=1)
    ax.spines['polar'].set_color('#D9E3E9')
    if title:
        ax.set_title(title, pad=18, fontsize=12)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()


def _details(scores: Mapping[str, Mapping[str, Any]], axes: Mapping[str, Any] | None, messages: Mapping[str, str] | None):
    rows = []
    for bid, data in scores.items():
        responses = list(data.get('reponses') or [])
        strengths = [_clean(r.get('reponse_label') or r.get('question')) for r in responses if r.get('reponse_val') == 2]
        gaps = [_clean(r.get('reponse_label') or r.get('question')) for r in responses if r.get('reponse_val') in (0, 1)]
        rows.append({
            'id': bid,
            'title': _clean(data.get('titre')) or bid,
            'pct': data.get('pct'),
            'code': _clean(data.get('capacite_code')),
            'label': _axis_label(axes, 'capacite', _clean(data.get('capacite_code'))),
            'strength': strengths[0] if strengths else _msg(messages, 'capacity_strength_none'),
            'gap': gaps[0] if gaps else _msg(messages, 'capacity_gap_none'),
        })
    return rows



def _capacity_groups(rows, axes: Mapping[str, Any] | None):
    # Regroupement piloté par l'axe capacité du classeur de paramétrage.
    axis = (axes or {}).get('capacite') or {}
    by_code: Dict[str, list] = {}
    for row in rows:
        by_code.setdefault(_clean(row.get('code')), []).append(row)
    groups = []
    for code, cfg in axis.items():
        code = _clean(code)
        group_rows = by_code.pop(code, [])
        if code == 'NON_EVALUEE' and not group_rows:
            continue
        try:
            order = int(float((cfg or {}).get('ordre', 9999) or 9999))
        except (TypeError, ValueError):
            order = 9999
        groups.append({
            'code': code,
            'label': _clean((cfg or {}).get('label')) or code,
            'color': _clean((cfg or {}).get('couleur')) or '#49697D',
            'description': _clean((cfg or {}).get('description')),
            'order': order,
            'rows': group_rows,
        })
    for code, group_rows in by_code.items():
        groups.append({'code': code, 'label': code, 'color': '#49697D', 'description': '', 'order': 9999, 'rows': group_rows})
    groups.sort(key=lambda g: (g['order'], g['label']))
    return groups


def _safe_hex_color(value: str, fallback: str = '#49697D') -> str:
    import re
    value = _clean(value)
    return value if re.fullmatch(r'#[0-9A-Fa-f]{6}', value or '') else fallback


def generate_autoevaluation_reports(
    out_html: Path,
    out_pdf: Path,
    scores: Mapping[str, Mapping[str, Any]],
    axes_config: Mapping[str, Any] | None,
    messages: Mapping[str, str] | None,
    meta: Mapping[str, Any] | None = None,
) -> None:
    meta = dict(meta or {})
    rows = _details(scores, axes_config, messages)
    groups = _capacity_groups(rows, axes_config)
    title = _msg(messages, 'autoeval_title')
    subtitle = _msg(messages, 'autoeval_subtitle')
    summary_title = _msg(messages, 'autoeval_summary_title')
    details_title = _msg(messages, 'autoeval_details_title')
    radar_title = _msg(messages, 'autoeval_radar_title')
    disclaimer = _msg(messages, 'autoeval_disclaimer')
    strength_label = _msg(messages, 'autoeval_strength_label')
    gap_label = _msg(messages, 'autoeval_gap_label')
    blocks_label = _msg(messages, 'autoeval_blocks_label')
    radar = _radar_png(scores, radar_title, meta.get('engine_params') or {})
    radar_uri = 'data:image/png;base64,' + base64.b64encode(radar).decode('ascii') if radar else ''

    identity = [(k, v) for k, v in [
        (_msg(messages, 'autoeval_identity_company_label'), meta.get('entreprise')),
        (_msg(messages, 'autoeval_identity_lot_label'), meta.get('lot')),
        (_msg(messages, 'autoeval_identity_address_label'), meta.get('adresse')),
        (_msg(messages, 'autoeval_identity_email_label'), meta.get('email')),
        (_msg(messages, 'autoeval_identity_analyst_label'), meta.get('analyste')),
        (_msg(messages, 'autoeval_identity_date_label'), meta.get('date')),
    ] if _clean(k) and _clean(v)]
    identity_html = ''.join(
        f'<div class="identity-line"><span>{html.escape(k)}</span><b>{html.escape(_clean(v))}</b></div>'
        for k, v in identity
    )

    evaluated_groups = [g for g in groups if g.get('code') != 'NON_EVALUEE']
    summary_cards = ''.join(
        f'<div class="metric" style="--metric-color:{html.escape(_safe_hex_color(g["color"]))}">'
        f'<b>{len(g["rows"])}</b><span>{html.escape(g["label"])}</span></div>'
        for g in evaluated_groups
    )

    group_sections = []
    for group in groups:
        if not group['rows']:
            continue
        group_color = _safe_hex_color(group['color'])
        block_html = []
        for r in group['rows']:
            block_html.append(
                '<article class="ae-block">'
                '<div class="ae-meta">'
                f'<strong class="block-code">{html.escape(r["id"])}</strong>'
                f'<h3>{html.escape(r["title"])}</h3>'
                f'<div class="meta-item"><span>{html.escape(_msg(messages, "autoeval_col_capacity_label"))}</span>'
                f'<b>{html.escape(r["label"])} - {html.escape(str(r["pct"]))} %</b></div>'
                '</div>'
                '<div class="ae-cell">'
                f'<div class="ae-cell-head">{html.escape(strength_label)}</div>'
                f'<p>{html.escape(r["strength"])}</p></div>'
                '<div class="ae-cell">'
                f'<div class="ae-cell-head">{html.escape(gap_label)}</div>'
                f'<p>{html.escape(r["gap"])}</p></div>'
                '</article>'
            )
        description_html = f'<p>{html.escape(group["description"])}</p>' if group['description'] else ''
        group_sections.append(
            f'<section class="capacity-group" style="--group-color:{html.escape(group_color)}">'
            '<div class="capacity-group-head">'
            f'<div><h2>{html.escape(group["label"])}</h2>{description_html}</div>'
            f'<strong>{len(group["rows"])} {html.escape(blocks_label)}</strong>'
            '</div>'
            f'<div class="capacity-group-body">{"".join(block_html)}</div></section>'
        )
    details_html = ''.join(group_sections)

    brand_value = _clean(meta.get('marque') or meta.get('entreprise'))
    if meta.get('logo_data_uri'):
        brand_html = f'<img src="{html.escape(_clean(meta.get("logo_data_uri")))}" alt="Logo">'
    else:
        brand_html = html.escape(brand_value)

    html_doc = f'''<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)}</title><style>
:root{{--navy:#071C24;--blue:#1F5D8C;--slate:#49697D;--pale:#EAF3F8;--pale2:#F7FAFC;--magenta:#A1003D;--magenta-light:#9C5B45;--text:#17232B;--muted:#667985;--line:#C9D6DE;--green:#176B55;--amber:#A66A12;--white:#fff}}
*{{box-sizing:border-box}}body{{margin:0;background:#f2f5f7;color:var(--text);font:14px/1.48 Calibri,Carlito,"Segoe UI",Arial,sans-serif}}h1,h2,.metric b{{font-family:"Calibri Light",Calibri,Carlito,"Segoe UI",Arial,sans-serif}}.hero{{background:var(--navy);border-bottom:5px solid var(--magenta);padding:27px max(25px,calc((100vw - 1500px)/2));display:flex;align-items:center;gap:22px}}.brand{{font-size:29px;font-weight:800;color:var(--magenta-light);border-right:1px solid rgba(240,106,154,.45);padding-right:22px}}.brand img{{max-height:40px;max-width:160px;display:block;object-fit:contain}}.hero h1{{margin:0;color:var(--magenta-light);font-size:31px;line-height:1;font-weight:700}}.hero p{{margin:7px 0 0;color:#F5C8D8;font-weight:600}}.container{{max-width:1500px;margin:0 auto;padding:22px}}.identity-strip{{background:#fff;border:1px solid var(--line);border-left:4px solid var(--blue);padding:10px 14px;margin-bottom:14px;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:5px 20px}}.identity-line{{display:grid;grid-template-columns:145px 1fr;gap:8px;font-size:12px}}.identity-line span{{color:var(--muted);text-transform:uppercase;font-size:9.5px;font-weight:700}}.identity-line b{{color:var(--navy);font-weight:600}}.summary{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:16px}}.metric{{background:#fff;border:1px solid var(--line);border-left:5px solid var(--metric-color,var(--blue));padding:13px 16px}}.metric b{{display:block;font-size:27px;color:var(--navy);font-weight:700}}.metric span{{font-size:10px;text-transform:uppercase;color:var(--muted);font-weight:700}}.section-title{{margin:18px 0 10px;color:var(--navy);font-size:18px}}.radar{{background:#fff;border:1px solid var(--line);padding:14px;text-align:center;margin-bottom:16px}}.radar img{{max-width:760px;width:100%}}.capacity-group{{background:#fff;border:1px solid var(--line);border-left:5px solid var(--group-color,var(--blue));margin-bottom:18px}}.capacity-group-head{{display:flex;align-items:flex-start;justify-content:space-between;gap:18px;padding:13px 15px;background:var(--pale2);border-bottom:1px solid var(--line)}}.capacity-group-head h2{{margin:0;color:var(--navy);font-size:18px}}.capacity-group-head p{{margin:4px 0 0;color:var(--muted);font-size:12px}}.capacity-group-head>strong{{white-space:nowrap;color:var(--group-color,var(--blue));font-size:12px;text-transform:uppercase}}.capacity-group-body{{padding:0}}.ae-block{{display:grid;grid-template-columns:18% 41% 41%;border-top:3px solid var(--magenta)}}.ae-block:first-child{{border-top:0}}.ae-meta{{background:var(--pale);border-right:1px solid var(--blue);padding:14px}}.block-code{{display:block;color:var(--magenta);font-family:"Calibri Light",Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-size:23px;line-height:1}}.ae-meta h3{{margin:6px 0 18px;color:var(--navy);font-size:14px;line-height:1.35}}.meta-item{{padding:10px 0;border-top:1px solid var(--line)}}.meta-item span{{display:block;margin-bottom:4px;color:var(--muted);font-size:10px;text-transform:uppercase;font-weight:700}}.meta-item b{{display:block;color:var(--group-color,var(--blue));font-size:12px;line-height:1.35}}.ae-cell{{border-right:1px solid var(--line);padding:0 14px 14px}}.ae-cell:last-child{{border-right:0}}.ae-cell-head{{margin:0 -14px 12px;padding:10px 12px;background:var(--navy);color:#fff;font-size:10px;text-transform:uppercase;letter-spacing:.04em;font-weight:800}}.ae-cell p{{margin:0;line-height:1.5}}.notice{{background:#fff;border:1px solid var(--line);border-left:4px solid var(--magenta);padding:12px 14px;color:var(--muted);margin-top:16px}}@media(max-width:950px){{.hero{{flex-wrap:wrap}}.identity-strip,.summary{{grid-template-columns:1fr}}.identity-line{{grid-template-columns:120px 1fr}}.ae-block{{display:block;margin-bottom:12px}}.ae-meta{{border-right:0}}.ae-cell{{border-right:0;border-top:1px solid var(--line)}}.capacity-group-head{{display:block}}.capacity-group-head>strong{{display:block;margin-top:8px}}}}
</style></head><body><header class="hero"><div class="brand">{brand_html}</div><div><h1>{html.escape(title.upper())}</h1><p>{html.escape(subtitle)}</p></div></header><main class="container"><section class="identity-strip">{identity_html}</section><h2 class="section-title">{html.escape(summary_title)}</h2><section class="summary">{summary_cards}</section>{f'<section class="radar"><img src="{radar_uri}" alt="{html.escape(radar_title)}"></section>' if radar_uri else ''}<h2 class="section-title">{html.escape(details_title)}</h2>{details_html}<section class="notice">{html.escape(disclaimer)}</section></main></body></html>'''
    Path(out_html).write_text(html_doc, encoding='utf-8')

    from reporting_v2 import (
        FONT, FONT_BOLD, V4, V4_NAVY, V4_BLUE, V4_MAGENTA, V4_MUTED,
        V4_LINE, V4_PALE, V4_PALE_2, V4_WHITE, _v4_doc
    )
    body = ParagraphStyle('AEBody', parent=V4['body'])
    small = ParagraphStyle('AESmall', parent=V4['small'])
    table_head = ParagraphStyle('AETableHead', parent=V4['label'], textColor=V4_WHITE, fontName=FONT_BOLD)
    group_title_style = ParagraphStyle('AEGroupTitle', parent=V4['h2'], textColor=V4_NAVY, fontName=FONT_BOLD)
    group_desc_style = ParagraphStyle('AEGroupDesc', parent=V4['small'], textColor=V4_MUTED)
    brand_pdf = _clean(meta.get('marque') or meta.get('entreprise'))
    doc = _v4_doc(out_pdf, title, marque=brand_pdf, logo_bytes=meta.get('logo_bytes'), footer_text=_msg(messages, 'autoeval_pdf_footer'))
    story = [Spacer(1, 5*mm), Paragraph(html.escape(title.upper()), V4['cover_title']), Paragraph(html.escape(subtitle), V4['cover_sub']), Spacer(1, 5*mm)]
    if identity:
        identity_data = [[Paragraph(html.escape(k), V4['label']), Paragraph(html.escape(_clean(v)), body)] for k, v in identity]
        identity_table = Table(identity_data, colWidths=[45*mm, 125*mm])
        identity_table.setStyle(TableStyle([
            ('VALIGN',(0,0),(-1,-1),'TOP'),
            ('LINEBELOW',(0,0),(-1,-2),0.25,V4_LINE),
            ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),4),
            ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
        ]))
        story += [identity_table, Spacer(1, 5*mm)]

    story += [Paragraph(html.escape(summary_title.upper()), V4['h1'])]
    metric_cells = []
    metric_styles = []
    for idx, g in enumerate(evaluated_groups):
        color = colors.HexColor(_safe_hex_color(g['color']))
        metric_cells.append(Paragraph(
            f'<font size="18"><b>{len(g["rows"])}</b></font><br/><font size="7">{html.escape(g["label"].upper())}</font>',
            body,
        ))
        metric_styles.append(('LINEBEFORE',(idx,0),(idx,0),3,color))
    if metric_cells:
        widths = [170*mm / len(metric_cells)] * len(metric_cells)
        metric = Table([metric_cells], colWidths=widths)
        metric.setStyle(TableStyle([
            ('BOX',(0,0),(-1,-1),0.45,V4_LINE),('INNERGRID',(0,0),(-1,-1),0.25,V4_LINE),
            ('BACKGROUND',(0,0),(-1,-1),V4_WHITE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
            ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
        ] + metric_styles))
        story += [metric, Spacer(1, 4*mm)]
    if radar:
        story += [Image(io.BytesIO(radar), width=160*mm, height=125*mm), Spacer(1, 4*mm)]

    story.append(Paragraph(html.escape(details_title.upper()), V4['h1']))
    cap_label = _msg(messages, 'autoeval_col_capacity_label')
    for group in groups:
        if not group['rows']:
            continue
        group_color = colors.HexColor(_safe_hex_color(group['color']))
        group_header = Table([[
            Paragraph(html.escape(group['label']), group_title_style),
            Paragraph(f'{len(group["rows"])} {html.escape(blocks_label)}', V4['small'])
        ]], colWidths=[135*mm, 35*mm])
        group_header.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),V4_PALE_2),('LINEBEFORE',(0,0),(0,0),3,group_color),
            ('LINEBELOW',(0,0),(-1,-1),0.45,V4_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),
            ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
            ('ALIGN',(1,0),(1,0),'RIGHT'),
        ]))
        story += [group_header]
        if group['description']:
            story += [Paragraph(html.escape(group['description']), group_desc_style), Spacer(1, 2*mm)]
        for r in group['rows']:
            meta_cell = [
                Paragraph(f'<font color="#A1003D"><b>{html.escape(r["id"])}</b></font>', V4['h2']),
                Spacer(1, 2),
                Paragraph(html.escape(r['title']), body),
                Spacer(1, 5),
                Paragraph(html.escape(cap_label.upper()), V4['label']),
                Paragraph(f'<b>{html.escape(r["label"])} - {html.escape(str(r["pct"]))} %</b>', body),
            ]
            strengths = [
                Paragraph(html.escape(strength_label.upper()), table_head),
                Spacer(1, 5),
                Paragraph(html.escape(r['strength']), body),
            ]
            gaps = [
                Paragraph(html.escape(gap_label.upper()), table_head),
                Spacer(1, 5),
                Paragraph(html.escape(r['gap']), body),
            ]
            block_table = Table([[meta_cell, strengths, gaps]], colWidths=[42*mm, 64*mm, 64*mm])
            block_table.setStyle(TableStyle([
                ('VALIGN',(0,0),(-1,-1),'TOP'),('BACKGROUND',(0,0),(0,0),V4_PALE),
                ('BOX',(0,0),(-1,-1),0.35,V4_LINE),('INNERGRID',(0,0),(-1,-1),0.25,V4_LINE),
                ('LINEBEFORE',(0,0),(0,0),2,V4_MAGENTA),
                ('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),
                ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),7),
            ]))
            story += [block_table, Spacer(1, 3*mm)]
        story += [Spacer(1, 3*mm)]

    story += [Paragraph(html.escape(disclaimer), small)]
    doc.build(story)

