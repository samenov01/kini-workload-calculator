"""Audit every available local MUP and the supplied template without changing inputs."""
import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path

from kini.audit import inventory_plan, verify_coverage, write_audit, audit_formulas
from kini.model import read_plan
from webapp import inspect_file


def main():
    root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', default='verification')
    args = parser.parse_args()
    destination = (root / args.output).resolve()
    destination.mkdir(parents=True, exist_ok=True)
    config = json.loads((root / 'kini_config.json').read_text(encoding='utf-8-sig'))
    known = {hashlib.sha256((root / g['file']).read_bytes()).hexdigest(): g for g in config['groups']}
    seen, inventory, courses, issues, source_reports = set(), [], [], [], []
    paths = sorted((root / 'input').glob('*.xlsx')) + sorted((root / '.kini-app/uploads').glob('*/*.xlsx'))
    for path in paths:
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest in seen:
            continue
        seen.add(digest)
        metadata = inspect_file(path)
        for sheet in metadata['sheets']:
            if digest in known:
                group = deepcopy(known[digest])
            else:
                if not sheet['intake'] or not sheet['program']:
                    raise ValueError(f'{path.name}: для проверки нужны программа и год поступления из шапки.')
                # No load is generated for these extra plans; student counts
                # are irrelevant to their independent source coverage audit.
                group = dict(name=f"{sheet['program']}-{sheet['intake']} / {path.stem}",
                             program=sheet['program'], intake=sheet['intake'], students=0, language='KZ')
            group['sheet'] = sheet['sheet']
            raw = inventory_plan(path, group)
            parsed = read_plan(group, path, issues)
            coverage = verify_coverage(raw, parsed)
            if any(coverage[k] for k in ('missing', 'extra', 'duplicates')):
                raise AssertionError(f'{path.name}: {coverage}')
            inventory.extend(raw)
            courses.extend(parsed)
            source_reports.append(dict(file=path.name, sha256=digest, sheet=sheet['sheet'],
                intake=group['intake'], program=group['program'], coverage=coverage,
                formulas=audit_formulas(path)))
            print(f"{path.name}: {coverage['represented']}/{coverage['expected']}")
    audit = write_audit(destination, inventory, courses, root / config['template'], config)
    audit['source_checks'] = source_reports
    audit['issues'] = [i.__dict__ for i in issues]
    (destination / 'Полная проверка проекта.json').write_text(
        json.dumps(audit, ensure_ascii=False, indent=2), encoding='utf-8')
    lines = ['# Полная проверка исходных МУПов', '',
             'Повторные загрузки одинаковых книг исключены по SHA-256. Исходные файлы не изменялись. Дополнительные планы проверены без формирования нагрузки и без предположений о численности групп.', '',
             '| Файл | Год приёма из шапки | Покрытие строк/семестров | Формулы проверены | Расхождения кэша | Циклы | Неподдержанные формулы |',
             '|---|---:|---:|---:|---:|---:|---:|']
    for r in source_reports:
        f, c = r['formulas'], r['coverage']
        lines.append(f"| {r['file']} | {r['intake']} | {c['represented']}/{c['expected']} | {f['checked']}/{f['total']} | {len(f['mismatches'])} | {len(f['cycles'])} | {len(f['unsupported'])} |")
    lines += ['', 'Построчное сопоставление всех 62 предметных строк образца с МУПами — в «Проверка МУПов и образца.md». Если МУПа нужного набора нет, найденные дисциплины других наборов приведены только для сравнения; они не подменяют отсутствующий источник.', '',
              '## Несогласованности исходных данных', '']
    for issue in issues:
        if issue.kind in ('кредиты', 'итог МУПа', 'структура языкового блока', 'часы общего блока', 'нет названия', 'нет семестра'):
            lines.append(f'- {issue.source}:{issue.row} — {issue.message}')
    lines += ['', '## Расхождения сохранённых формул МУПов', '',
              '| Файл | Лист:ячейка | Формула | Кэш | По сохранённым входам |', '|---|---|---|---:|---:|']
    for source in source_reports:
        for m in source['formulas']['mismatches']:
            lines.append(f"| {source['file']} | {m['sheet']}:{m['cell']} | {m['formula']} | {m['cached']} | {m['from_cached_inputs']} |")
    (destination / 'Обзор проверки.md').write_text('\n'.join(lines) + '\n', encoding='utf-8')
    print('TOTAL', audit['coverage'])


if __name__ == '__main__':
    main()
