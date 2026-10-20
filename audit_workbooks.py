from pathlib import Path
from collections import Counter
import argparse, json, re, warnings, ast, operator
import openpyxl
from openpyxl.formula.translate import Translator
warnings.simplefilter('ignore',UserWarning)
parser = argparse.ArgumentParser(description='Audit cached formulas in local Excel workbooks.')
parser.add_argument('--source', type=Path, default=Path('input'))
src = parser.parse_args().source
Path('workbook_review').mkdir(exist_ok=True)
result=[]
ops={ast.Add:operator.add,ast.Sub:operator.sub,ast.Mult:operator.mul,ast.Div:operator.truediv,ast.UAdd:operator.pos,ast.USub:operator.neg}
def arithmetic(s):
    def visit(n):
        if isinstance(n,ast.Constant) and isinstance(n.value,(int,float)): return n.value
        if isinstance(n,ast.BinOp): return ops[type(n.op)](visit(n.left),visit(n.right))
        if isinstance(n,ast.UnaryOp): return ops[type(n.op)](visit(n.operand))
        raise ValueError(s)
    return visit(ast.parse(s,mode='eval').body)
for p in sorted(src.glob('*.xlsx')):
    wb=openpyxl.load_workbook(p); cv=openpyxl.load_workbook(p,data_only=True)
    report={'file':p.name,'sheets':[]}
    for ws in wb:
        patterns=Counter(); mismatch=[]; unsupported=[]; dependencies=Counter(); annotations=[]; checked=0
        for row in ws:
            for c in row:
                if c.comment: annotations.append({'cell':c.coordinate,'comment':c.comment.text})
                if c.data_type!='f': continue
                f=c.value
                for sh in re.findall(r"'([^']+)'!",f): dependencies[sh]+=1
                try: patterns[Translator(f,origin=c.coordinate).translate_formula('ZZ10000')]+=1
                except Exception: patterns[f]+=1
                def ref(match):
                    sh=(match.group(1) or match.group(2) or ws.title).replace("''", "'")
                    address=match.group(3).replace('$','')
                    val=cv[sh][address].value
                    if val is None or isinstance(val,str): val=0
                    return str(val)
                try:
                    expr=f[1:]
                    def sums(m):
                        a,b=m.group(1).replace('$','').split(':')
                        return str(sum(c.value for rr in cv[ws.title][a+':'+b] for c in rr if isinstance(c.value,(int,float))))
                    expr=re.sub(r'(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)',sums,expr,flags=re.I)
                    expr=re.sub(r"(?:'((?:[^']|'')+)'!|([\w ]+)!)?(\$?[A-Z]{1,3}\$?\d+)",ref,expr)
                    expr=re.sub(r'SUM\(([^()]*)\)',lambda m:'('+m.group(1).replace(',', '+')+')',expr,flags=re.I)
                    value=arithmetic(expr); old=cv[ws.title][c.coordinate].value
                    if isinstance(old,(int,float)):
                        checked+=1
                        if abs(value-old)>1e-7: mismatch.append({'cell':c.coordinate,'formula':f,'cached':old,'from_cached_inputs':value})
                except Exception: unsupported.append({'cell':c.coordinate,'formula':f})
        report['sheets'].append({'sheet':ws.title,'checked':checked,'mismatches':mismatch,'unsupported':unsupported,'dependencies':dict(dependencies),'formula_patterns':patterns.most_common(),'comments':annotations})
    result.append(report)
Path('workbook_review/audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
for b in result:
    print(b['file'])
    for s in b['sheets']:
        print(s['sheet'],'checked',s['checked'],'cache mismatches',len(s['mismatches']),'unsupported',len(s['unsupported']),'comments',len(s['comments']),'links',','.join(s['dependencies']))
        for m in s['mismatches'][:3]: print(m)
