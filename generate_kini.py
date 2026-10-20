"""Запуск: python generate_kini.py --config kini_config.json"""
from __future__ import annotations

import argparse
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import shutil
import sys

from kini.model import Issue, assign_keys, build_loads, read_plan
from kini.excel import write_book
from kini.audit import inventory_plan, verify_coverage, write_audit, write_kazakh_comparison


FILES = {
    'PE': 'МУП Process Engineering 2026 (1) (2).xlsx',
    'CS': 'МУП CS 26.xlsx',
    'LOG': 'муп логистика 26.xlsx',
    'MECH': 'муп мехатроника 26.xlsx',
    'EEE': 'МУП EEE 26.xlsx',
}
TEMPLATE = 'расчет KINI 2026-27 (10).xlsx'


def initialize(root, source):
    if (root / 'kini_config.json').exists():
        raise ValueError('kini_config.json уже существует. Повторная инициализация не изменяет настройки.')
    for filename in [*FILES.values(), TEMPLATE]:
        if not (source / filename).is_file():
            raise ValueError(f'Не найден файл {source / filename}')
    (root / 'input').mkdir(exist_ok=True)
    (root / 'template').mkdir(exist_ok=True)
    for filename in FILES.values():
        shutil.copy2(source / filename, root / 'input' / filename)
    shutil.copy2(source / TEMPLATE, root / 'template' / TEMPLATE)
    counts = {'PE': 4, 'CS': 8, 'LOG': 7, 'MECH': 7, 'EEE': 9}
    config = {
        'template': 'template/' + TEMPLATE,
        'output': 'output',
        'groups': [dict(name=f'{p}-26', program=p, intake=2026, students=counts[p], language='KZ',
                        file='input/' + filename, students_by_year={}, exclude_rows=[], overrides={})
                   for p, filename in FILES.items()],
        'aliases': {}, 'routing': {},
    }
    (root / 'kini_config.json').write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding='utf-8')


def run(config_path, progress=None, output_dir=None):
    root = config_path.resolve().parent
    config = json.loads(config_path.read_text(encoding='utf-8-sig'))
    template = (root / config['template']).resolve()
    output = Path(output_dir).resolve() if output_dir else (root / config.get('output', 'output')).resolve()
    if output in [template.parent] or any(output == (root / g['file']).resolve().parent for g in config['groups']):
        raise ValueError('Папка результатов должна отличаться от папок исходных МУПов и шаблона.')
    if not config['groups']:
        raise ValueError('Добавьте хотя бы одну группу в kini_config.json.')
    names = [g['name'] for g in config['groups']]
    if len(set(names)) != len(names):
        raise ValueError('Имена групп должны быть уникальны.')
    for g in config['groups']:
        if not isinstance(g.get('students'), int) or g['students'] < 0:
            raise ValueError(f'{g["name"]}: students должен быть целым неотрицательным числом.')
        if not isinstance(g.get('intake'), int) or not 2000 <= g['intake'] <= 2200:
            raise ValueError(f'{g["name"]}: intake — год поступления, например 2026.')
        if g.get('language', 'KZ') not in ('KZ', 'RU', 'ENG', 'DE'):
            raise ValueError(f'{g["name"]}: language должен быть KZ, RU, ENG или DE.')
    courses, issues, sources, inventory = [], [], [], []
    for g in config['groups']:
        if progress:
            progress(f'Чтение МУПа: {g["name"]}')
        path = (root / g['file']).resolve()
        if not path.is_file():
            raise ValueError(f'Не найден МУП: {path}')
        entries = read_plan(g, path, issues)
        source_inventory = inventory_plan(path, g)
        # Явные пользовательские исключения остаются видны в замечаниях;
        # автоматические исключения парсера не допускаются.
        for item in source_inventory:
            if item['row'] in g.get('exclude_rows', []):
                item['semesters'] = {}
            else:
                item['semesters'] = {s: v for s, v in item['semesters'].items()
                    if not g.get('overrides', {}).get(f'{item["row"]}:{s}', {}).get('exclude')}
        inventory.extend(source_inventory)
        for c in entries:
            count = g.get('students_by_year', {}).get(str(c.year), g['students'])
            if not isinstance(count, int) or count < 0:
                raise ValueError(f'{g["name"]}: некорректный контингент на {c.year}.')
            c.students = count
            if c.kind == 'ordinary' and not c.control.strip():
                issues.append(Issue('форма контроля', c.group, c.file, c.row,
                                    f'{c.name}: форма контроля в МУПе не заполнена; экзамен рассчитан по заданному правилу 0,1 × студенты.', c.year))
        courses.extend(entries)
        sources.append(dict(group=g['name'], path=str(path), sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
                            contributions=len(entries)))
    assign_keys(courses, config.get('aliases', {}))
    coverage = verify_coverage(inventory, courses)
    if any(coverage[k] for k in ('missing', 'extra', 'duplicates')):
        raise ValueError('Проверка полноты МУПов не пройдена: ' + json.dumps(coverage, ensure_ascii=False))
    loads, main, hourly = build_loads(courses, template, config, issues)
    # Конфигурация направления не может распределять одни и те же часы дважды.
    years = sorted({c.year for c in courses})
    output.mkdir(parents=True, exist_ok=True)
    results = []
    for year in years:
        if progress:
            progress(f'Формирование {year}–{year+1}')
        path = output / f'KINI_НАГРУЗКА_{year}-{year+1}.xlsx'
        selected = [l for l in loads if l.year == year]
        results.append(write_book(template, path, year, selected, main, hourly, issues))
        print(f'{path.name}: {len(selected)} строк нагрузки; I семестр {results[-1]["totals"]["1"]}, '
              f'II семестр {results[-1]["totals"]["2"]}; контрольные разности = 0.')
    audit = write_audit(output, inventory, courses, template, config)
    write_kazakh_comparison(output, audit['template_rows'], results, config)
    manifest = dict(config=str(config_path.resolve()), template_sha256=hashlib.sha256(template.read_bytes()).hexdigest(),
                    sources=sources, results=results, issues=[asdict(i) for i in issues],
                    courses=[asdict(c) for c in courses], coverage=coverage,
                    template_checks=dict(rows=len(audit['template_rows']), formulas=audit['template_formulas']))
    (output / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    report = ['# Результаты формирования KINI', '',
              'Расчёт предварительный: незаданные часы оставлены пустыми и отмечены жёлтым. '
              'Итоги включают только заполненные часы; это не утверждённая полная потребность.', '',
              '| Учебный год | Строк нагрузки | I семестр | II семестр | Год |',
              '|---|---:|---:|---:|---:|']
    for r in results:
        t = r['totals']
        report.append(f'| {r["year"]}–{r["year"]+1} | {r["load_rows"]} | {t["1"]} | {t["2"]} | {t["0"]} |')
    report += ['', 'Для каждого файла проверены СРСП, экзамен, суммы строк, семестров, года, '
               'сводов и распределения. Все контрольные разности H:T равны нулю.', '',
               'В каждой книге только лист «Расчет». Отдельные итоговые листы временно отключены. '
               'Нераспределённые часы сохранены отдельным блоком на листе «Расчет». '
               'Назначения по направлениям берутся только из однозначных соответствий шаблона.', '',
               f'Независимая проверка полноты: {coverage["represented"]} из {coverage["expected"]} исходных записей «группа / строка / семестр». Потерь и повторов нет. '
               'Составляющие общих блоков показаны отдельными строками без повторного начисления общего бюджета. '
               'Подробная сверка с образцом — в «Проверка МУПов и образца.md».', '',
               'Подробности по исходным строкам и замечания сохранены в manifest.json; источники также указаны в примечаниях ячеек.']
    (output / 'Результаты.md').write_text('\n'.join(report), encoding='utf-8')
    return manifest


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description='Сформировать расчёты нагрузки KINI по учебным годам.')
    parser.add_argument('--config', type=Path, default=Path('kini_config.json'))
    parser.add_argument('--output', type=Path, help='Папка результатов вместо заданной в конфигурации.')
    parser.add_argument('--init', action='store_true', help='Создать настройки и скопировать исходные МУПы 2026.')
    parser.add_argument('--source', type=Path, default=Path.home() / 'Downloads' / 'Telegram Desktop')
    args = parser.parse_args()
    try:
        if args.init:
            initialize(args.config.resolve().parent, args.source)
        run(args.config, output_dir=args.output)
    except (ValueError, FileNotFoundError, PermissionError) as exc:
        print(f'Ошибка: {exc}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
