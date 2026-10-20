"""Распознавание таблиц по заголовкам, независимо от имени файла и сдвигов."""
from dataclasses import dataclass, asdict
import ast
import operator
import re


@dataclass
class Layout:
    header: int
    name: int
    code: int
    credits: int
    control: int
    hours: tuple
    semesters: dict


def detect_layout(ws):
    candidate = None
    for row in ws.iter_rows(max_row=min(200, ws.max_row), max_col=min(200, ws.max_column)):
        semesters = {}
        for c in row:
            text = str(c.value or '').strip().lower()
            match = re.match(r'^(\d{1,2})\s*[.\-–]?\s*(сем|term|semester)', text)
            if match:
                sem = int(match[1])
                if 1 <= sem <= 16:
                    if sem in semesters:
                        raise ValueError(f'Лист «{ws.title}»: повторяется заголовок семестра {sem}.')
                    semesters[sem] = c.column
        if len(semesters) >= 2 and 1 in semesters:
            candidate = row[0].row, semesters
            break
    if not candidate:
        raise ValueError(f'Лист «{ws.title}»: не найдена таблица с заголовками «1 семестр», «2 семестр» и далее.')
    header, semesters = candidate
    if sorted(semesters) != list(range(1, max(semesters) + 1)):
        raise ValueError(f'Лист «{ws.title}»: в заголовках есть пропуски семестров.')
    cells = [c for row in ws.iter_rows(min_row=max(1, header-5), max_row=header,
                                      max_col=min(semesters.values())-1) for c in row if c.value]

    def find(label, predicate):
        matches = {c.column for c in cells if predicate(str(c.value).lower().replace('\n', ' '))}
        if len(matches) != 1:
            raise ValueError(f'Лист «{ws.title}»: не удалось однозначно определить столбец «{label}».')
        return matches.pop()

    primary_names = [c for c in cells if any(x in str(c.value).lower() for x in ('наименование дисциплин', 'пән аталуы'))]
    name = find('Наименование дисциплины', lambda s: any(x in s for x in
                (('наименование дисциплин', 'пән аталуы') if primary_names else ('name of the discipline',))))
    code = find('Код дисциплины', lambda s: any(x in s for x in ('код дисциплин', 'code of discipline', 'сode of discipline', 'пән коды')))
    credits = find('Кредиты', lambda s: ('ects' in s or 'кредит' in s) and
                   not any(x in s for x in ('семестр', 'semester', 'распределени', 'бөліну')))
    control = find('Экзамен', lambda s: any(x in s for x in ('экзамен', 'емтихан', 'exam')) and
                   not any(x in s for x in ('распределение', 'форм контроля', 'distribution')))
    lecture = find('Лекции', lambda s: bool(re.search(r'\bлк\b|лекци|lecture', s)))
    practical = find('Практические', lambda s: bool(re.search(r'\bпз\b|практическ|practical', s)))
    lab = find('Лабораторные', lambda s: bool(re.search(r'\bлз\b|лаборатор|laboratory', s)))
    return Layout(header, name, code, credits, control, (lecture, practical, lab), semesters)


def numeric(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value
    if isinstance(value, str):
        text = value.strip().replace('\xa0', '').replace(' ', '').replace(',', '.')
        if re.fullmatch(r'[+-]?\d+(\.\d+)?', text):
            return float(text)
    return None


class PlanValues:
    """Сохранённые значения; для новых файлов без кэша — простые формулы Excel."""
    def __init__(self, cached, formulas):
        self.cached, self.formulas = cached, formulas
        self.memo, self.active = {}, set()

    def get(self, sheet, address):
        key = sheet, address
        c = self.cached[sheet][address]
        f = self.formulas[sheet][address]
        if c.value is not None or f.data_type != 'f':
            return c.value
        if key in self.memo:
            return self.memo[key]
        if key in self.active:
            raise ValueError(f'Циклическая формула {sheet}!{address}. Исправьте МУП.')
        self.active.add(key)
        try:
            def cell(match):
                sh = (match[1] or sheet).replace("''", "'")
                value = self.get(sh, match[2].replace('$', ''))
                return repr(numeric(value) or 0)
            def summed(match):
                total = 0
                for row in self.formulas[sheet][match[1].replace('$', '')]:
                    total += sum(numeric(self.get(sheet, c.coordinate)) or 0 for c in row)
                return str(total)
            expr = f.value[1:].strip()
            expr = re.sub(r'(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)', summed, expr)
            expr = re.sub(r"(?:'((?:[^']|'')+)'!)?(\$?[A-Z]+\$?\d+)", cell, expr)
            expr = re.sub(r'SUM\(([^()]*)\)', lambda m: '(' + m[1].replace(',', '+').rstrip('+') + ')', expr, flags=re.I)
            ops = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv}
            def calculate(n):
                if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
                    return n.value
                if isinstance(n, ast.BinOp): return ops[type(n.op)](calculate(n.left), calculate(n.right))
                if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.UAdd, ast.USub)):
                    return calculate(n.operand) * (-1 if isinstance(n.op, ast.USub) else 1)
                raise ValueError('unsupported')
            result = calculate(ast.parse(expr, mode='eval').body)
            self.memo[key] = result
            return result
        except (SyntaxError, KeyError, TypeError, ZeroDivisionError, ValueError) as exc:
            raise ValueError(f'Не удалось вычислить {sheet}!{address}: {f.value}. Откройте МУП в Excel, пересчитайте и сохраните.') from exc
        finally:
            self.active.remove(key)
